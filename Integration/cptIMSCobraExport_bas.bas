Attribute VB_Name = "cptIMSCobraExport_bas"
'<cpt_version>v3.4.70</cpt_version>
Option Explicit
Private destFolder As String
Private BCWSxport As Boolean
Private BCWPxport As Boolean
Private ETCxport As Boolean
Private WhatIfxport As Boolean 'v3.2
Private ResourceLoaded As Boolean
Private ExportMilestones As Boolean 'v3.3.13
Private MasterProject As Boolean
Private ACTfilename As String
Private RESfilename As String
Private BCR_WP() As String
Private BCR_ID As String
Private BCRxport As Boolean
Private BCR_Error As Boolean
Private fProject, fCAID1, fCAID1t, fCAID3, fCAID3t, fWP, fCAM, fPCNT, fAssignPcnt, fEVT, fCAID2, fCAID2t, fMilestone, fMilestoneWeight, fBCR, fWhatIf, fResID As String 'v3.3.0, v3.4.3
Private dateFmt As String 'v3.3.5
Private CustTextFields() As String
Private EntFields() As String
Private CustNumFields() As String
Private CustOLCodeFields() As String
Private ActFound As Boolean
Private CAID3_Used As Boolean
Private CAID2_Used As Boolean
Private Milestones_Used As Boolean
Private AssignmentPCNT_Used As Boolean 'v3.3.2
Private TimeScaleExport As Boolean
Private TsvScale As String 'v3.4
Private DescExport As Boolean
Private ErrMsg As String
Private WPDescArray() As WP_Descriptions
Private WPDescCount As Integer
Private ActIDCounter As Integer 'v3.3.5
Private subprojectIDs As Boolean 'v3.4.3
Private Type WP_Descriptions
    WP_ID As String
    Desc As String
End Type
Private Type ACTrowWP
    SubProject As String 'v3.4.3
    CAID1 As String
    CAID3 As String
    CAID2 As String
    Desc As String
    CAM As String
    WP As String
    Resource As String
    ID As String
    ShortID As String 'v3.3.5
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
    FStart As String 'v3.3.0
    FFinish As String 'v3.3.0
    BWork As Double
    BCost As Double
    FWork As Double 'v3.3.0
    FCost As Double 'v3.3.0
    AssignmentBStart As String
    AssignmentBFinish As String
    AssignmentFStart As String 'v3.3.0
    AssignmentFFinish As String 'v3.3.0
    AssignmentBWork As Double
    AssignmentBCost As Double
    AssignmentFWork As Double 'v3.3.0
    AssignmentFCost As Double 'v3.3.0
    AssignmentTSVWork As Double 'v3.3.2
    AssignmentTSVCost As Double 'v3.3.2
    EVT As String
    MSID As String
    MSWeight As String
    AssignmentCount As Integer
End Type
Private noFolderSelected As Boolean

Sub Export_IMS()

    Dim xportFrm As cptIMSCobraExport_frm
    Dim xportFormat As String
    Dim curProj As Project
    Dim i As Integer

    On Error GoTo CleanUp

    Set curProj = ActiveProject

    curProj.Application.Calculation = pjManual
    curProj.Application.DisplayAlerts = False

    If curProj.Subprojects.Count > 0 And InStr(curProj.FullName, "<>") > 0 And curProj.ReadOnly <> True Then
        MsgBox "Master Project Files with Subprojects must be opened Read Only"
        GoTo Quick_Exit
    End If

    If curProj.Subprojects.Count > 0 Then
        MasterProject = True
    Else
        MasterProject = False
    End If

    ReadCustomFields curProj

    Set xportFrm = New cptIMSCobraExport_frm

    With xportFrm

        On Error Resume Next

        .resBox.List = Split("Name,Code,Initials", ",")

        'populate listboxes
        Dim vArray As Variant
        vArray = Split(Join(CustTextFields, ",") & "," & Join(CustOLCodeFields, ",") & "," & Join(EntFields, ","), ",") 'v3.3.9
        If vArray(UBound(vArray)) = "" Then ReDim Preserve vArray(UBound(vArray) - 1) 'v3.3.10
        Call exportQuickSort(vArray, 0, UBound(vArray)) 'v3.4.7
        .caID1Box.List = Split("WBS," & Join(vArray, ","), ",")
        .caID2Box.List = Split("<None>," & Join(vArray, ","), ",")
        .caID3Box.List = Split("<None>," & Join(vArray, ","), ",")
        .wpBox.List = vArray
        .camBox.List = Split("Contact," & Join(vArray, ","), ",")
        .evtBox.List = vArray
        .mswBox.List = Split("<None>,BaselineWork,BaselineCost,Work,Cost," & Join(CustNumFields, ",") & "," & Join(vArray, ","), ",") 'v3.3.9
        .bcrBox.List = Split("<None>," & Join(vArray, ","), ",")
        .projBox.List = Split("<None>," & Join(vArray, ","), ",") 'v3.4.3
        .whatifBox.List = Split("<None>," & Join(vArray, ","), ",")
        vArray = Split(Join(CustTextFields, ",") & "," & Join(CustNumFields, ",") & "," & Join(CustOLCodeFields, ",") & "," & Join(EntFields, ","), ",") 'v3.3.9
        If vArray(UBound(vArray)) = "" Then ReDim Preserve vArray(UBound(vArray) - 1) 'v3.3.10
        Call exportQuickSort(vArray, 0, UBound(vArray)) 'v3.4.7
        .msidBox.List = Split("<None>,UniqueID," & Join(vArray, ","), ",")
        Call exportQuickSort(CustNumFields, 1, UBound(CustNumFields)) 'v3.4.7
        .PercentBox.List = Split("Physical % Complete,% Complete," & Join(CustNumFields, ","), ",")
        .AsgnPcntBox.List = Split("<None>," & Join(CustNumFields, ","), ",")
        .DateFormat_Combobox.List = Split("M/D/YYYY,D/M/YYYY", ",") 'v3.3.5
        .WeekStartCombobox.List = Split("Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", ",")
        .WeekStartCombobox.ListIndex = curProj.StartWeekOn - 1
        .ScaleCombobox.List = Split("Weekly,Monthly", ",") 'v3.4
        
        On Error GoTo CleanUp
        ErrMsg = "Please try again, or contact the developer if this message repeats."
        '********************************************
        'On Error GoTo 0 '**Used for Debugging ONLY**
        '********************************************

        .Show

        If .Tag = "Cancel" Then
            Set xportFrm = Nothing
            Set curProj = Nothing
            Exit Sub
        End If

        If .Tag = "DataCheck" Then
            CAID3_Used = .CAID3TxtBox.Enabled
            CAID2_Used = .CAID2TxtBox.Enabled
            DataChecks curProj
            Set xportFrm = Nothing
            GoTo Quick_Exit
        End If

        If .MPPBtn.Value = True Then
            Set xportFrm = Nothing
            xportFormat = "MPP"
        ElseIf .XMLBtn.Value = True Then
            Set xportFrm = Nothing
            xportFormat = "XML"
        ElseIf .CSVBtn.Value = True Then
            BCWSxport = .BCWS_Checkbox.Value
            BCWPxport = .BCWP_Checkbox.Value
            ETCxport = .ETC_Checkbox.Value
            BCRxport = .BcrBtn.Value
            ExportMilestones = .Milestone_CheckBox.Value
            WhatIfxport = .WhatIf_CheckBox.Value 'v3.2
            BCR_ID = .BCR_ID_TextBox
            ResourceLoaded = .ResExportCheckbox
            TimeScaleExport = .exportTPhaseCheckBox
            TsvScale = .ScaleCombobox.Value 'v3.4
            Set xportFrm = Nothing
            xportFormat = "CSV"
            CAID3_Used = .CAID3TxtBox.Enabled
            CAID2_Used = .CAID2TxtBox.Enabled
            If .msidBox.Value = "<None>" Or .mswBox.Value = "<None>" Or .msidBox.Value = "" Or .mswBox.Value = "" Then
                Milestones_Used = False
            Else
                Milestones_Used = True
            End If
            If .projBox.Value = "<None>" Or .projBox.Value = "" Then 'v3.4.3
                subprojectIDs = False
            Else
                subprojectIDs = True
            End If
            If .AsgnPcntBox = "<None>" Or .AsgnPcntBox = "" Then
                AssignmentPCNT_Used = False
            Else
                AssignmentPCNT_Used = True
            End If
            DescExport = .exportDescCheckBox.Value
            dateFmt = .DateFormat_Combobox.Value
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
    If noFolderSelected = False Then MsgBox "An error was encountered." & vbCr & vbCr & ErrMsg 'v3.4.2
    Exit Sub

Quick_Exit:

    curProj.Application.Calculation = pjAutomatic
    curProj.Application.DisplayAlerts = True
    Set curProj = Nothing

    Exit Sub

End Sub

Private Function get_assignment_timescalevalues(ByVal tAss As Assignment) As Double 'v3.3.2
    
    Dim tsvs As TimeScaleValues
    Dim tsv As TimeScaleValue
    Dim tempTotal As Double
    
    If tAss.Resource.Type = pjResourceTypeCost Then
        Set tsvs = tAss.TimeScaleData(tAss.BaselineStart, tAss.BaselineFinish, pjAssignmentTimescaledBaselineCost)
    Else
        Set tsvs = tAss.TimeScaleData(tAss.BaselineStart, tAss.BaselineFinish, pjAssignmentTimescaledBaselineWork)
    End If
    
    tempTotal = 0
    
    For Each tsv In tsvs
    
        If tsv.Value <> "" Then
            
            tempTotal = tempTotal + tsv.Value
        
        End If
        
    Next
    
    get_assignment_timescalevalues = tempTotal

End Function

Private Sub DataChecks(ByVal curProj As Project)

    Dim WPChecks() As WPDataCheck
    Dim wpFound As Boolean
    Dim CAMChecks() As CAMDataCheck
    Dim CAfound As Boolean
    Dim TaskChecks() As TaskDataCheck
    Dim taskFound As Boolean
    Dim t As Task
    Dim tAss As Assignment
    Dim tasses As Assignments
    Dim tAssBStart As String
    Dim tAssBFin As String
    Dim tAssBWork As String
    Dim tempID As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim wpCount As Integer
    Dim camCount As Integer
    Dim taskCount As Integer
    Dim X As Integer
    Dim i As Integer
    Dim errorStr As String
    Dim ErrorCounter As Integer
    Dim tempBValue As Double
    Dim tempBWork As Double
    Dim tempFWork As Double 'v3.3.0
    Dim tempTSVWork As Double 'v3.3.2

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
    If CAID2_Used = True Then
        fCAID2 = docProps("fCAID2").Value
        fCAID2t = docProps("fCAID2t").Value
    End If
    If Milestones_Used = True Then 'v3.2.6
        fMilestone = docProps("fMSID").Value
        fMilestoneWeight = docProps("fMSW").Value
    End If
    fPCNT = docProps("fPCNT").Value

    destFolder = SetDirectory(curProj.ProjectSummaryTask.Project)

    taskCount = 0
    taskFound = False

    '**Scan Task Data**

    If curProj.Subprojects.Count > 0 Then

        Set subProjs = curProj.Subprojects

        For Each subProj In subProjs

            FileOpen Name:=subProj.Path, ReadOnly:=True

            Set curSProj = ActiveProject

            For Each t In curSProj.Tasks

                If Not t Is Nothing Then

                    If t.Active = True And t.ExternalTask = False Then

                        taskCount = taskCount + 1
                        taskFound = True
                        ReDim Preserve TaskChecks(1 To taskCount)

                        With TaskChecks(taskCount)

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
                            If Milestones_Used = True Then
                                .MSID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                .MSWeight = t.GetField(FieldNameToFieldConstant(fMilestoneWeight))
                            End If
                            If t.GetField(pjTaskBaselineWork) <> "" Then 'v3.4.6
                                .BWork = t.BaselineWork / 60 'v3.3.2
                            Else
                                .BWork = 0
                            End If
                            If t.GetField(pjTaskBaselineCost) <> "" Then 'v3.4.6
                                .BCost = t.BaselineCost
                            Else
                                .BCost = 0
                            End If
                            .FWork = t.Work / 60 'v3.3.2
                            .FCost = t.Cost 'v3.3.0
                            .CAM = t.GetField(FieldNameToFieldConstant(fCAM))
                            .AssignmentBStart = "NA"
                            .AssignmentBFinish = "NA"
                            .AssignmentFStart = "NA" 'v3.3.0
                            .AssignmentFFinish = "NA" 'v3.3.0
                            .AssignmentBCost = 0
                            .AssignmentBWork = 0
                            .AssignmentFCost = 0 'v3.3.0
                            .AssignmentFWork = 0 'v3.3.0
                            .AssignmentTSVWork = 0 'v3.3.2
                            .AssignmentTSVCost = 0 'v3.3.2
                            .BStart = t.BaselineStart
                            .BFinish = t.BaselineFinish
                            .FStart = t.Start 'v3.3.0
                            .FFinish = t.Finish 'v3.3.0

                            Set tasses = t.Assignments
                            .AssignmentCount = tasses.Count

                            For Each tAss In tasses

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
                                
                                If .AssignmentFStart = "NA" Then 'v3.3.0
                                    .AssignmentFStart = tAss.Start
                                Else
                                    If tAss.Start < .AssignmentFStart Then 'v3.3.14
                                        .AssignmentFStart = tAss.Start
                                    End If
                                End If

                                If .AssignmentFFinish = "NA" Then 'v3.3.0
                                    .AssignmentFFinish = tAss.Finish
                                Else
                                    If tAss.Finish > .AssignmentFFinish Then
                                        .AssignmentFFinish = tAss.Finish
                                    End If
                                End If

                                .AssignmentBCost = .AssignmentBCost + tAss.BaselineCost
                                If tAss.ResourceType <> pjResourceTypeWork Then 'v3.3.2
                                    .AssignmentTSVCost = .AssignmentTSVCost + get_assignment_timescalevalues(tAss)
                                End If
                                .AssignmentFCost = .AssignmentFCost + tAss.Cost 'v3.3.0
                                
                                If tAss.BaselineWork = "" Or tAss.ResourceType <> pjResourceTypeWork Then 'v3.2.1
                                    tempBWork = 0
                                    tempTSVWork = 0 'v3.3.2
                                Else
                                    tempBWork = tAss.BaselineWork
                                    tempTSVWork = get_assignment_timescalevalues(tAss) 'v3.3.2
                                End If
                                .AssignmentBWork = .AssignmentBWork + tempBWork / 60 'v3.3.2
                                .AssignmentTSVWork = .AssignmentTSVWork + tempTSVWork / 60 'v3.3.2
                                
                                If tAss.Work = "" Or tAss.ResourceType <> pjResourceTypeWork Then 'v3.3.0
                                    tempFWork = 0
                                Else
                                    tempFWork = tAss.Work
                                End If
                                .AssignmentFWork = .AssignmentFWork + tempFWork / 60 'v3.3.2

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

                If t.Active = True And t.ExternalTask = False Then

                    taskCount = taskCount + 1
                    taskFound = True
                    ReDim Preserve TaskChecks(1 To taskCount)

                    With TaskChecks(taskCount)

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
                            If Milestones_Used = True Then
                                .MSID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                .MSWeight = t.GetField(FieldNameToFieldConstant(fMilestoneWeight))
                            End If
                            If t.GetField(pjTaskBaselineWork) <> "" Then 'v3.4.6
                                .BWork = t.BaselineWork / 60 'v3.3.2
                            Else
                                .BWork = 0
                            End If
                            If t.GetField(pjTaskBaselineCost) <> "" Then 'v3.4.6
                                .BCost = t.BaselineCost
                            Else
                                .BCost = 0
                            End If
                            .FWork = t.Work / 60 'v3.3.2
                            .FCost = t.Cost 'v3.3.0
                            .CAM = t.GetField(FieldNameToFieldConstant(fCAM))
                            .AssignmentBStart = "NA"
                            .AssignmentBFinish = "NA"
                            .AssignmentFStart = "NA" 'v3.3.0
                            .AssignmentFFinish = "NA" 'v3.3.0
                            .AssignmentBCost = 0
                            .AssignmentBWork = 0
                            .AssignmentFCost = 0 'v3.3.0
                            .AssignmentFWork = 0 'v3.3.0
                            .AssignmentTSVWork = 0 'v3.3.2
                            .AssignmentTSVCost = 0 'v3.3.2
                            .BStart = t.BaselineStart
                            .BFinish = t.BaselineFinish
                            .FStart = t.Start 'v3.3.0
                            .FFinish = t.Finish 'v3.3.0

                            Set tasses = t.Assignments
                            .AssignmentCount = tasses.Count

                            For Each tAss In tasses

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
                                
                                If .AssignmentFStart = "NA" Then 'v3.3.0
                                    .AssignmentFStart = tAss.Start
                                Else
                                    If tAss.Start < .AssignmentFStart Then
                                        .AssignmentFStart = tAss.Start
                                    End If
                                End If

                                If .AssignmentFFinish = "NA" Then 'v3.3.0
                                    .AssignmentFFinish = tAss.Finish
                                Else
                                    If tAss.Finish > .AssignmentFFinish Then
                                        .AssignmentFFinish = tAss.Finish
                                    End If
                                End If

                                .AssignmentBCost = .AssignmentBCost + tAss.BaselineCost
                                If tAss.ResourceType <> pjResourceTypeWork Then 'v3.3.2
                                    .AssignmentTSVCost = .AssignmentTSVCost + get_assignment_timescalevalues(tAss)
                                End If
                                .AssignmentFCost = .AssignmentFCost + tAss.Cost 'v3.3.0
                                
                                If tAss.BaselineWork = "" Or tAss.ResourceType <> pjResourceTypeWork Then 'v3.2.1
                                    tempBWork = 0
                                    tempTSVWork = 0 'v3.3.2
                                Else
                                    tempBWork = tAss.BaselineWork
                                    tempTSVWork = get_assignment_timescalevalues(tAss) 'v3.3.2
                                End If
                                .AssignmentBWork = .AssignmentBWork + tempBWork / 60 'v3.3.2
                                .AssignmentTSVWork = .AssignmentTSVWork + tempTSVWork / 60 'v3.3.2
                                
                                If tAss.Work = "" Or tAss.ResourceType <> pjResourceTypeWork Then 'v3.3.0
                                    tempFWork = 0
                                Else
                                    tempFWork = tAss.Work
                                End If
                                .AssignmentFWork = .AssignmentFWork + tempFWork / 60 'v3.3.2

                            Next tAss

                    End With

                End If

            End If

        Next t

    End If

    ACTfilename = destFolder & "\DataChecks_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

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

    For X = 1 To taskCount

        If CAID3_Used = True And CAID2_Used = True Then
            tempID = TaskChecks(X).CAID1 & "/" & TaskChecks(X).CAID2 & "/" & TaskChecks(X).CAID3
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            tempID = TaskChecks(X).CAID1 & "/" & TaskChecks(X).CAID2
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            tempID = TaskChecks(X).CAID1
        End If

        If TaskChecks(X).CAM <> "" And TaskChecks(X).WP <> "" Then

            CAfound = False

            If camCount = 0 Then

                camCount = 1

                ReDim CAMChecks(1 To camCount)

                With CAMChecks(camCount)

                    .ID_str = tempID
                    .CAM_Test = TaskChecks(X).CAM
                    .CAM_Error = False

                End With

            Else

                For i = 1 To camCount

                    If CAMChecks(i).ID_str = tempID Then

                        CAfound = True

                        If TaskChecks(X).CAM <> CAMChecks(i).CAM_Test Then
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
                        .CAM_Test = TaskChecks(X).CAM
                        .CAM_Error = False

                    End With

                End If

            End If

        End If

        If TaskChecks(X).WP <> "" Then

            wpFound = False

            If wpCount = 0 Then

                wpCount = 1

                ReDim WPChecks(1 To wpCount)

                With WPChecks(wpCount)

                    .ID_Test = tempID
                    .WP_ID = TaskChecks(X).WP
                    .EVT_Test = TaskChecks(X).EVT
                    .WP_DupError = False
                    .EVT_Error = False

                End With

            Else

                For i = 1 To wpCount

                    If WPChecks(i).WP_ID = TaskChecks(X).WP Then

                        wpFound = True

                        If tempID <> WPChecks(i).ID_Test Then

                            WPChecks(i).WP_DupError = True

                        End If

                        If TaskChecks(X).EVT <> WPChecks(i).EVT_Test Then

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
                        .WP_ID = TaskChecks(X).WP
                        .EVT_Test = TaskChecks(X).EVT
                        .WP_DupError = False
                        .EVT_Error = False

                    End With

                End If

            End If

        End If

next_task:

        '**Report Tasks Missing Metadata**

        If TaskChecks(X).WP <> "" Or (TaskChecks(X).EVT <> "" And TaskChecks(X).EVT <> "NA" And TaskChecks(X).EVT <> "N/A") Or TaskChecks(X).BCost <> 0 Or TaskChecks(X).BWork <> 0 Then 'v3.2.2

            If TaskChecks(X).BWork = 0 Then tempBValue = TaskChecks(X).BCost Else tempBValue = TaskChecks(X).BWork

            If TaskChecks(X).WP = "" Or TaskChecks(X).EVT = "" Or TaskChecks(X).BStart = "NA" Or TaskChecks(X).BFinish = "NA" Or tempBValue = 0 Then

                ErrorCounter = ErrorCounter + 1

                If CAID3_Used = True And CAID2_Used = True Then

                    With TaskChecks(X)

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

                    With TaskChecks(X)

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

                    With TaskChecks(X)

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

    Next X

    Print #1, vbCrLf & "Total Task Errors Found: " & ErrorCounter

    '**Report Multiple CAM Assignments**

    Print #1, vbCrLf & vbCrLf & "CAM Errors - The following items reflect multiple CAM assignments per Control Account as interpreted by Cobra (based on a unique CA record ID constructed from Concatenated CA ID Values"

    Print #1, vbCrLf & "CA ID String"

    ErrorCounter = 0

    For X = 1 To camCount

        If CAMChecks(X).CAM_Error = True Then

            ErrorCounter = ErrorCounter + 1

            Print #1, CAMChecks(X).ID_str

        End If

    Next X

    Print #1, vbCrLf & "Total CAM Errors Found: " & ErrorCounter

    '**Report Duplicate WP IDs & Multiple EVTs**

    Print #1, vbCrLf & vbCrLf & "Work Package Errors - The following Work Package IDs are duplicated across multiple CA ID values and/or are assigned multiple EVTs"

    Print #1, vbCrLf & "Work Package,Duplicate WP ID,Multiple EVTs"

    ErrorCounter = 0

    For X = 1 To wpCount

        If WPChecks(X).WP_DupError = True Or WPChecks(X).EVT_Error = True Then

            ErrorCounter = ErrorCounter + 1

            With WPChecks(X)
                errorStr = .WP_ID & "," & .WP_DupError & "," & .EVT_Error
            End With

            Print #1, errorStr

            errorStr = ""
        End If

    Next X

    Print #1, vbCrLf & "Total Work Package Errors Found: " & ErrorCounter

    '**Reporting Assignment Baseline Issues (Values and Dates)**

    Print #1, vbCrLf & vbCrLf & "Task Assignment Baseline Discrepancies - The following Tasks have vertical traceability errors with their Assignment Baseline Values and/or Baseline Dates. Note that tasks with mixed elements of cost (labor & material/cost) that include labor rates may result in false positives." 'v3.3.0

    Print #1, vbCrLf & "UID,Task Baseline Work,Assignment Baseline Work,Assignment TimeScaled Work,Task Baseline Cost,Assignment Baseline Cost,Assignment TimeScaled Cost,Task Baseline Start,Assignment Baseline Start,Task Baseline Finish,Assignment Baseline Finish, Assignment Count" 'v3.3.2

    ErrorCounter = 0

    For X = 1 To taskCount

        With TaskChecks(X)

            If .AssignmentCount > 0 Then 'v3.2.3

                If Round(CDec(.AssignmentTSVWork), 2) <> Round(CDec(.AssignmentBWork), 2) Or (Round(CDec(.AssignmentTSVCost), 2) <> Round(CDec(.AssignmentBCost), 2) And .AssignmentBWork = 0) Or Round(CDec(.BCost), 2) <> Round(CDec(.AssignmentBCost), 2) Or Round(CDec(.BWork), 2) <> Round(CDec(.AssignmentBWork), 2) Or .BStart <> .AssignmentBStart Or .BFinish <> .AssignmentBFinish Then 'v3.3.2

                    ErrorCounter = ErrorCounter + 1

                    errorStr = .UID & ","
                    errorStr = errorStr & .BWork & ","
                    errorStr = errorStr & .AssignmentBWork & ","
                    errorStr = errorStr & .AssignmentTSVWork & "," 'v3.3.2
                    errorStr = errorStr & .BCost & ","
                    errorStr = errorStr & .AssignmentBCost & ","
                    errorStr = errorStr & .AssignmentTSVCost & "," 'v3.3.2
                    errorStr = errorStr & .BStart & ","
                    errorStr = errorStr & .AssignmentBStart & ","
                    errorStr = errorStr & .BFinish & ","
                    errorStr = errorStr & .AssignmentBFinish & "," 'v3.2.3
                    errorStr = errorStr & .AssignmentCount 'v3.2.3

                    Print #1, errorStr
                    errorStr = ""

                End If

            End If
            
            If .AssignmentCount = 0 And (.BCost <> 0 Or .BWork <> 0) Then 'v3.2.3
            
                ErrorCounter = ErrorCounter + 1
                
                errorStr = .UID & ","
                errorStr = errorStr & .BWork & ","
                errorStr = errorStr & .AssignmentBWork & ","
                errorStr = errorStr & .AssignmentTSVWork & "," 'v3.3.2
                errorStr = errorStr & .BCost & ","
                errorStr = errorStr & .AssignmentBCost & ","
                errorStr = errorStr & .BStart & ","
                errorStr = errorStr & .AssignmentBStart & ","
                errorStr = errorStr & .AssignmentTSVCost & "," 'v3.3.2
                errorStr = errorStr & .BFinish & ","
                errorStr = errorStr & .AssignmentBFinish & "," 'v3.2.3
                errorStr = errorStr & .AssignmentCount 'v3.2.3
            
                Print #1, errorStr
                errorStr = ""
            
            End If

        End With

    Next X

    Print #1, vbCrLf & "Total Task Assignment Baseline Errors Found: " & ErrorCounter 'v3.3.0

    '**Reporting Assignment Forecast Issues (Values and Dates)**

    Print #1, vbCrLf & vbCrLf & "Task Assignment Forecast Discrepancies - The following Tasks have potential vertical traceability errors with their Assignment Forecast Values and/or Forecast Dates"

    Print #1, vbCrLf & "UID,Task Work,Assignment Work,Task Cost,Assignment Cost,Task Start,Assignment Start,Task Finish,Assignment Finish, Assignment Count"

    ErrorCounter = 0

    For X = 1 To taskCount

        With TaskChecks(X)

            If .AssignmentCount > 0 Then

                If Round(CDec(.FCost), 2) <> Round(CDec(.AssignmentFCost), 2) Or Round(CDec(.FWork), 2) <> Round(CDec(.AssignmentFWork), 2) Or .FStart <> .AssignmentFStart Or .FFinish <> .AssignmentFFinish Then

                    ErrorCounter = ErrorCounter + 1

                    errorStr = .UID & ","
                    errorStr = errorStr & .FWork & ","
                    errorStr = errorStr & .AssignmentFWork & ","
                    errorStr = errorStr & .FCost & ","
                    errorStr = errorStr & .AssignmentFCost & ","
                    errorStr = errorStr & .FStart & ","
                    errorStr = errorStr & .AssignmentFStart & ","
                    errorStr = errorStr & .FFinish & ","
                    errorStr = errorStr & .AssignmentFFinish & ","
                    errorStr = errorStr & .AssignmentCount

                    Print #1, errorStr
                    errorStr = ""

                End If

            End If
            
            If .AssignmentCount = 0 And (.FCost <> 0 Or .FWork <> 0) Then
            
                ErrorCounter = ErrorCounter + 1
                
                errorStr = .UID & ","
                errorStr = errorStr & .FWork & ","
                errorStr = errorStr & .AssignmentFWork & ","
                errorStr = errorStr & .FCost & ","
                errorStr = errorStr & .AssignmentFCost & ","
                errorStr = errorStr & .FStart & ","
                errorStr = errorStr & .AssignmentFStart & ","
                errorStr = errorStr & .FFinish & ","
                errorStr = errorStr & .AssignmentFFinish & ","
                errorStr = errorStr & .AssignmentCount
            
                Print #1, errorStr
                errorStr = ""
            
            End If

        End With

    Next X
    
    Print #1, vbCrLf & "Total Task Assignment Forecast Errors Found: " & ErrorCounter

    MsgBox "Data Check Report saved to " & destFolder

    Shell "explorer.exe" & " " & destFolder, vbNormalFocus

    Close #1

End Sub

Private Sub MPP_Export(ByVal curProj As Project)

    Dim subProj As SubProject
    Dim subProjs As Subprojects

    destFolder = SetDirectory(curProj.ProjectSummaryTask.Project)

    If curProj.Subprojects.Count > 0 Then

        Set subProjs = curProj.Subprojects

        For Each subProj In subProjs

            subProj.SourceProject.SaveAs Name:=destFolder & "\" & subProj.SourceProject.Name
            curProj.Subprojects(subProj.Index).SourceProject = destFolder & "\" & subProj.SourceProject.Name

        Next subProj

        curProj.SaveAs Name:=destFolder & "\" & curProj.ProjectSummaryTask.Project

    Else

        curProj.SaveAs Name:=destFolder & "\" & curProj.ProjectSummaryTask.Project

    End If

End Sub
Private Sub XML_Export(ByVal curProj As Project)

    Dim subProj As SubProject
    Dim subProjs As Subprojects

    destFolder = SetDirectory(curProj.ProjectSummaryTask.Project)

    If curProj.Subprojects.Count > 0 Then

        Set subProjs = curProj.Subprojects

        For Each subProj In subProjs

            subProj.SourceProject.SaveAs Name:=destFolder & "\" & subProj.SourceProject.Name, FormatID:="MSProject.XML"

        Next subProj


    Else

        curProj.SaveAs Name:=destFolder & "\" & curProj.ProjectSummaryTask.Project, FormatID:="MSProject.XML"

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
    If WhatIfxport = True Then 'v3.2
        fWhatIf = docProps("fWhatIf").Value
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
    If Milestones_Used Then
        fMilestone = docProps("fMSID").Value
        fMilestoneWeight = docProps("fMSW").Value
    End If
    fPCNT = docProps("fPCNT").Value
    If AssignmentPCNT_Used = True Then 'v3.3.2
        fAssignPcnt = docProps("fAssignPcnt").Value 'v3.3.0
    End If
    fResID = docProps("fResID").Value
    If subprojectIDs Then fProject = docProps("fProject").Value 'v3.4.3, v3.4.4

    BCR_Error = False

    destFolder = SetDirectory(curProj.ProjectSummaryTask.Project)

    '*******************
    '****BCR Review*****
    '*******************

    If (BCWSxport = True Or WhatIfxport = True) And BCRxport = True Then 'v3.3.15
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
    
    '*******************
    '**What-if Export***
    '*******************

    If WhatIfxport = True Then 'v3.2

        WhatIf_Export curProj

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
    Dim ProjID, CAID1, CAID3, WP, CAM, EVT, UID, CAID2, ResName, MSWeight, ID, PCNT As String 'v3.3.0, v3.4.3
    Dim Milestone As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim X As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String
    Dim tempID As String 'v3.3.3
    Dim headerStr As String 'v3.4.3
    Dim outputStr As String 'v3.4.3

    If ResourceLoaded = False Then

        ACTfilename = destFolder & "\BCWP ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        
        'v3.4.3 - refactored header output code
        headerStr = ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete,"
        
        If CAID3_Used = True And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & "," & fCAID3t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            headerStr = fCAID1t & headerStr
        End If
        
        If subprojectIDs Then 'v3.4.3
            headerStr = "Project," & headerStr
        End If
        
        Print #1, headerStr

        X = 1
        ActFound = False

        If curProj.Subprojects.Count > 0 Then

            Set subProjs = curProj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.Active = True And t.ExternalTask = False Then

                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                                If subprojectIDs Then 'v3.4.3
                                    ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                                End If
                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                If Milestones_Used = True Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                End If
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

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    err.Raise 1
                                End If
                                
                                'v3.4.3 - refactored data output code
                                
                                If EVT = "B" Or EVT = "N" Or EVT = "B Milestone" Or EVT = "N Earning Rules" Then

                                    outputStr = WP & "," & UID & "," & Format(t.Start, dateFmt) & "," & Format(t.Finish, dateFmt) & "," & Format(t.ActualStart, dateFmt) & "," & Format(t.ActualFinish, dateFmt) & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) & ","

                                    If CAID3_Used = True And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        outputStr = CAID1 & "," & outputStr
                                    End If
                                    
                                    If subprojectIDs Then 'v3.4.3
                                        outputStr = ProjID & "," & outputStr
                                    End If
                                    
                                    Print #1, outputStr

                                ElseIf EVT = "C" Or EVT = "C % Work Complete" Then

                                    'store ACT info
                                    'WP Data
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(X).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(X).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = t.Finish
                                        ACTarray(X).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        If subprojectIDs Then 'v3.4.3
                                            ACTarray(X).SubProject = ProjID
                                        End If
                                        ACTarray(X).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = t.ActualFinish
                                        End If
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(X).sumBCWS = 1
                                            ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        Else
                                            ACTarray(X).sumBCWS = 1
                                            ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        End If
                                        ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                        X = X + 1
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
                                    ReDim Preserve ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(i).sumBCWS = 1
                                        ACTarray(i).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(i).sumBCWS = 1
                                        ACTarray(i).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
                                    ActFound = True

                                ElseIf EVT = "E" Or EVT = "F" Or EVT = "G" Or EVT = "H" Or EVT = "E 50/50" Or EVT = "F 0/100" Or _
                                    EVT = "G 100/0" Or EVT = "H User Defined" Or EVT = "A" Or EVT = "A Level of Effort" Or EVT = "O" Or EVT = "O Earned As Spent" Then '3.4.4

                                    'store ACT info
                                    'WP Data
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(X).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(X).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = t.Finish
                                        ACTarray(X).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        If subprojectIDs Then 'v3.4.3
                                            ACTarray(X).SubProject = ProjID
                                        End If
                                        ACTarray(X).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = t.ActualFinish
                                        End If

                                        X = X + 1
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
                                    ReDim Preserve ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
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

                    If t.Active = True And t.ExternalTask = False Then

                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If subprojectIDs Then 'v3.4.3
                                ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                            End If
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                            End If
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

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                err.Raise 1
                            End If

                            'v3.4.3 - refactored data output code

                            If EVT = "B" Or EVT = "B Milestone" Or EVT = "N" Or EVT = "N Earning Rules" Then

                                outputStr = WP & "," & UID & "," & Format(t.Start, dateFmt) & "," & Format(t.Finish, dateFmt) & "," & Format(t.ActualStart, dateFmt) & "," & Format(t.ActualFinish, dateFmt) & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) & ","

                                If CAID3_Used = True And CAID2_Used = True Then
                                    outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    outputStr = CAID1 & "," & outputStr
                                End If
                                
                                If subprojectIDs Then 'v3.4.3
                                    outputStr = ProjID & "," & outputStr
                                End If
                                
                                Print #1, outputStr

                            ElseIf EVT = "C" Or EVT = "C % Work Complete" Then

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(X).sumBCWS = 1
                                        ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(X).sumBCWS = 1
                                        ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
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
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If
                                If t.BaselineWork <> 0 Then
                                    ACTarray(X).sumBCWS = 1
                                    ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                Else
                                    ACTarray(X).sumBCWS = 1
                                    ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                End If
                                ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                X = X + 1
                                ActFound = True

                            ElseIf EVT = "E" Or EVT = "E 50/50" Or EVT = "F" Or EVT = "F 0/100" Or EVT = "G" Or EVT = "G 100/0" Or _
                                EVT = "H" Or EVT = "H User Defined" Or EVT = "A" Or EVT = "A Level of Effort" Or EVT = "O" Or EVT = "O Earned As Spent" Then '3.4.4

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
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
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
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

                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, dateFmt)
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, dateFmt)

                'v3.4.3 - refactored data output code
                
                outputStr = ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, dateFmt) & "," & Format(ACTarray(i).FFinish, dateFmt) & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog & ","

                If CAID3_Used = True And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAID3 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    outputStr = ACTarray(i).CAID1 & "," & outputStr
                End If
                
                If subprojectIDs Then 'v3.4.3
                    outputStr = ACTarray(i).SubProject & "," & outputStr
                End If
                
                Print #1, outputStr

            Next i
        End If

        Close #1

    Else '**Resource Loaded**

        ACTfilename = destFolder & "\BCWP ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        
        'v3.4.3 - refactored header output code
        headerStr = ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete,Resource,"
        
        If CAID3_Used = True And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & "," & fCAID3t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            headerStr = fCAID1t & headerStr
        End If
        
        If subprojectIDs Then 'v3.4.3
            headerStr = "Project," & headerStr
        End If
        
        Print #1, headerStr

        X = 1
        ActFound = False

        If curProj.Subprojects.Count > 0 Then

            Set subProjs = curProj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.Active = True And t.ExternalTask = False Then

                            If t.BaselineWork > 0 Or t.BaselineCost > 0 Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If subprojectIDs Then 'v3.4.3
                                    ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                                End If
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                If Milestones_Used = True Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                End If
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
                                ResName = "" 'v3.3.0

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    err.Raise 1
                                End If
                                
                                'v3.4.3 - refactored data output code

                                If EVT = "B" Or EVT = "B Milestone" Or EVT = "N" Or EVT = "N Earned Rules" Then

                                    outputStr = WP & "," & UID & "," & Format(t.Start, dateFmt) & "," & Format(t.Finish, dateFmt) & "," & Format(t.ActualStart, dateFmt) & "," & Format(t.ActualFinish, dateFmt) & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) & ","

                                    If CAID3_Used = True And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        outputStr = CAID1 & "," & outputStr
                                    End If
                                    
                                    If subprojectIDs Then 'v3.4.3
                                        outputStr = ProjID & "," & outputStr
                                    End If
                                    
                                    Print #1, outputStr
    
                                ElseIf EVT = "L" Or EVT = "L Assignment % Complete" Then 'v3.3.0
                                
                                    'store ACT info
                                    'WP Data
                                    
                                    Set tAss = t.Assignments
                                        
                                    For Each tAssign In tAss
                                    
                                        ResName = tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource))
                                        tempID = ID & "/" & ResName
                                        
                                        If X = 1 Then
    
                                            'create new WP line in ACTarrray
                                            ReDim ACTarray(1 To X)
                                            If CAID3_Used = True Then
                                                ACTarray(X).CAID3 = CAID3
                                            End If
                                            ACTarray(X).CAM = CAM
                                            ACTarray(X).Resource = ResName
                                            ACTarray(X).ID = tempID
                                            ACTarray(X).CAID1 = CAID1
                                            ACTarray(X).EVT = EVT
                                            If CAID2_Used = True Then
                                                ACTarray(X).CAID2 = CAID2
                                            End If
                                            If subprojectIDs Then 'v3.4.3
                                                ACTarray(X).SubProject = ProjID
                                            End If
                                            ACTarray(X).WP = WP
                                            ACTarray(X).FFinish = tAssign.Finish
                                            ACTarray(X).FStart = tAssign.Start
                                            If tAssign.ActualStart <> "NA" Then ACTarray(X).AStart = tAssign.ActualStart
                                            If tAssign.ActualFinish <> "NA" Then ACTarray(X).AFinish = tAssign.ActualFinish
                                            
                                            If tAssign.BaselineWork <> 0 Then
                                                ACTarray(X).sumBCWS = tAssign.BaselineWork / 60
                                                ACTarray(X).sumBCWP = tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                            Else
                                                ACTarray(X).sumBCWS = tAssign.BaselineCost
                                                ACTarray(X).sumBCWP = tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                            End If
                                            
                                            If ACTarray(X).sumBCWS > 0 Then ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100
    
                                            X = X + 1
                                            ActFound = True
    
                                            GoTo Next_Assign_A
    
                                        End If
    
                                        For i = 1 To UBound(ACTarray)
                                            If ACTarray(i).ID = tempID Then
                                                'Found an existing matching WP line
                                                If ACTarray(i).FStart > tAssign.Start Then
                                                    ACTarray(i).FStart = tAssign.Start
                                                End If
                                                If ACTarray(i).FFinish < tAssign.Finish Then
                                                    ACTarray(i).FFinish = tAssign.Finish
                                                End If
                                                If tAssign.ActualStart <> "NA" Then
                                                    If ACTarray(i).AStart = 0 Then
                                                        ACTarray(i).AStart = tAssign.ActualStart
                                                    Else
                                                        If tAssign.ActualStart < ACTarray(i).AStart Then
                                                            ACTarray(i).AStart = tAssign.ActualStart
                                                        End If
                                                    End If
                                                End If
                                                If tAssign.ActualFinish <> "NA" Then
                                                    If ACTarray(i).AFinish = 0 Then
                                                        ACTarray(i).AFinish = tAssign.ActualFinish
                                                    Else
                                                        If tAssign.ActualFinish > ACTarray(i).AFinish Then
                                                            ACTarray(i).AFinish = tAssign.ActualFinish
                                                        End If
                                                    End If
                                                End If
                                                If tAssign.BaselineWork <> 0 Then
                                                    ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + tAssign.BaselineWork / 60
                                                    ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100)
                                                Else
                                                    ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + tAssign.BaselineCost
                                                    ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100)
                                                End If
    
                                                If ACTarray(i).sumBCWS > 0 Then ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100
    
                                                GoTo Next_Assign_A
                                            End If
                                        Next i
    
                                        'No match found, create new WP line in ACTarrray
                                        ReDim Preserve ACTarray(1 To X)
                                        
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).Resource = ResName
                                        ACTarray(X).ID = tempID
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = tAssign.Finish
                                        ACTarray(X).FStart = tAssign.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        If subprojectIDs Then 'v3.4.3
                                            ACTarray(X).SubProject = ProjID
                                        End If
                                        ACTarray(X).WP = WP
                                        If tAssign.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = tAssign.ActualStart
                                        End If
                                        If tAssign.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = tAssign.ActualFinish
                                        End If
                                        If tAssign.BaselineWork <> 0 Then
                                            ACTarray(X).sumBCWS = tAssign.BaselineWork / 60
                                            ACTarray(X).sumBCWP = tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                        Else
                                            ACTarray(X).sumBCWS = tAssign.BaselineCost
                                            ACTarray(X).sumBCWP = tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                        End If
                                        
                                        If ACTarray(X).sumBCWS > 0 Then ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                        X = X + 1
                                        
                                        ActFound = True
                                        
Next_Assign_A:

                                    Next tAssign
                                
                                ElseIf EVT = "C" Or EVT = "C % Work Complete" Then

                                    'store ACT info
                                    'WP Data
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(X).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(X).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = t.Finish
                                        ACTarray(X).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        If subprojectIDs Then 'v3.4.3
                                            ACTarray(X).SubProject = ProjID
                                        End If
                                        ACTarray(X).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = t.ActualFinish
                                        End If
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(X).sumBCWS = t.BaselineWork / 60
                                            ACTarray(X).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        Else
                                            ACTarray(X).sumBCWS = t.BaselineCost
                                            ACTarray(X).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        End If
                                        ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                        X = X + 1
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
                                    ReDim Preserve ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(X).sumBCWS = t.BaselineWork / 60 'v3.3.0
                                        ACTarray(X).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100 'v3.3.0
                                    Else
                                        ACTarray(X).sumBCWS = t.BaselineCost
                                        ACTarray(X).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100 'v3.3.0
                                    End If
                                    ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
                                    ActFound = True

                                ElseIf EVT = "E" Or EVT = "E 50/50" Or EVT = "F" Or EVT = "F 0/100" Or EVT = "G" Or EVT = "G 100/0" Or _
                                    EVT = "H" Or EVT = "H User Defined" Or EVT = "A" Or EVT = "A Level of Effort" Or EVT = "O" Or EVT = "O Earned As Spent" Then '3.4.4

                                    'store ACT info
                                    'WP Data
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(X).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(X).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = t.Finish
                                        ACTarray(X).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        If subprojectIDs Then 'v3.4.3
                                            ACTarray(X).SubProject = ProjID
                                        End If
                                        ACTarray(X).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = t.ActualFinish
                                        End If

                                        X = X + 1
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
                                    ReDim Preserve ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
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

                    If t.Active = True And t.ExternalTask = False Then

                        If t.BaselineWork > 0 Or t.BaselineCost > 0 Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If subprojectIDs Then 'v3.4.3
                                ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                            End If
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                            End If
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
                            ResName = "" 'v3.3.0

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                err.Raise 1
                            End If

                            If EVT = "B" Or EVT = "B Milestone" Or EVT = "N" Or EVT = "N Earned Rules" Then

                                'v3.4.3 - refactored data output code
                                
                                outputStr = WP & "," & UID & "," & Format(t.Start, dateFmt) & "," & Format(t.Finish, dateFmt) & "," & Format(t.ActualStart, dateFmt) & "," & Format(t.ActualFinish, dateFmt) & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) & ","

                                If CAID3_Used = True And CAID2_Used = True Then
                                    outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    outputStr = CAID1 & "," & outputStr
                                End If
                                
                                If subprojectIDs Then 'v3.4.3
                                    outputStr = ProjID & "," & outputStr
                                End If
                                
                                Print #1, outputStr
                                
                            ElseIf EVT = "L" Or EVT = "L Assignment % Complete" Then 'v3.3.0
                                
                                'store ACT info
                                'WP Data
                                
                                Set tAss = t.Assignments
                                    
                                For Each tAssign In tAss
                                
                                    ResName = tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource))
                                    tempID = ID & "/" & ResName
                                    
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).Resource = ResName
                                        ACTarray(X).ID = tempID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        If subprojectIDs Then 'v3.4.3
                                            ACTarray(X).SubProject = ProjID
                                        End If
                                        ACTarray(X).WP = WP
                                        ACTarray(X).FFinish = tAssign.Finish
                                        ACTarray(X).FStart = tAssign.Start
                                        If tAssign.ActualStart <> "NA" Then ACTarray(X).AStart = tAssign.ActualStart
                                        If tAssign.ActualFinish <> "NA" Then ACTarray(X).AFinish = tAssign.ActualFinish
                                        
                                        If tAssign.BaselineWork <> 0 Then
                                            ACTarray(X).sumBCWS = tAssign.BaselineWork / 60
                                            ACTarray(X).sumBCWP = tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                        Else
                                            ACTarray(X).sumBCWS = tAssign.BaselineCost
                                            ACTarray(X).sumBCWP = tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                        End If
                                        
                                        If ACTarray(X).sumBCWS > 0 Then ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                        X = X + 1
                                        
                                        ActFound = True

                                        GoTo Next_Assign_B

                                    End If

                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = tempID Then
                                            'Found an existing matching WP line
                                            If ACTarray(i).FStart > tAssign.Start Then
                                                ACTarray(i).FStart = tAssign.Start
                                            End If
                                            If ACTarray(i).FFinish < tAssign.Finish Then
                                                ACTarray(i).FFinish = tAssign.Finish
                                            End If
                                            If tAssign.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = tAssign.ActualStart
                                                Else
                                                    If tAssign.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = tAssign.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If tAssign.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = tAssign.ActualFinish
                                                Else
                                                    If tAssign.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = tAssign.ActualFinish
                                                    End If
                                                End If
                                            End If
                                            If tAssign.BaselineWork <> 0 Then
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + tAssign.BaselineWork / 60
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100)
                                            Else
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + tAssign.BaselineCost
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100)
                                            End If

                                            If ACTarray(i).sumBCWS > 0 Then ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100

                                            GoTo Next_Assign_B
                                            
                                        End If
                                    Next i

                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To X)
                                    
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).Resource = ResName
                                    ACTarray(X).ID = tempID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = tAssign.Finish
                                    ACTarray(X).FStart = tAssign.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If tAssign.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = tAssign.ActualStart
                                    End If
                                    If tAssign.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = tAssign.ActualFinish
                                    End If
                                    If tAssign.BaselineWork <> 0 Then
                                        ACTarray(X).sumBCWS = tAssign.BaselineWork / 60
                                        ACTarray(X).sumBCWP = tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                    Else
                                        ACTarray(X).sumBCWS = tAssign.BaselineCost
                                        ACTarray(X).sumBCWP = tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                    End If
                                    
                                    If ACTarray(X).sumBCWS > 0 Then ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
                                    
                                    ActFound = True
                                    
Next_Assign_B:
                                
                                Next tAssign

                            ElseIf EVT = "C" Or EVT = "C % Work Complete" Then

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(X).sumBCWS = t.BaselineWork / 60
                                        ACTarray(X).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(X).sumBCWS = t.BaselineCost
                                        ACTarray(X).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
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
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If
                                If t.BaselineWork <> 0 Then
                                    ACTarray(X).sumBCWS = t.BaselineWork / 60
                                    ACTarray(X).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                Else
                                    ACTarray(X).sumBCWS = t.BaselineCost
                                    ACTarray(X).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                End If
                                ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                X = X + 1
                                ActFound = True

                            ElseIf EVT = "E" Or EVT = "E 50/50" Or EVT = "F" Or EVT = "F 0/100" Or EVT = "G" Or EVT = "G 100/0" Or _
                                EVT = "H" Or EVT = "H User Defined" Or EVT = "A" Or EVT = "A Level of Effort" Or EVT = "O" Or EVT = "O Earned As Spent" Then '3.4.4

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
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
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
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

                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, dateFmt)
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, dateFmt)

                'v3.4.3 - refactored data output code
                
                outputStr = ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, dateFmt) & "," & Format(ACTarray(i).FFinish, dateFmt) & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog & "," & ACTarray(i).Resource & ","

                If CAID3_Used = True And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAID3 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    outputStr = ACTarray(i).CAID1 & "," & outputStr
                End If
                
                If subprojectIDs Then 'v3.4.3
                    outputStr = ACTarray(i).SubProject & "," & outputStr
                End If
                
                Print #1, outputStr

            Next i
        End If

        Close #1

    End If

End Sub

Private Sub ETC_Export(ByVal curProj As Project)

    Dim t As Task
    Dim tAss As Assignments
    Dim tAssign As Assignment
    Dim ProjID, CAID1, CAID3, WP, CAM, EVT, UID, CAID2, MSWeight, ID, PCNT, ShortID As String 'v3.3.5, v3.4.3
    Dim Milestone As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim X As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String
    Dim headerStr As String 'v3.4.3
    Dim outputStr As String 'v3.4.3

    '*******************
    '****ETC Export****
    '*******************
    
    ActIDCounter = 0 'v3.3.5

    If ResourceLoaded = False Then

        ACTfilename = destFolder & "\ETC ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        
        'v3.4.3 - refactored header output code
        headerStr = ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date,"
        
        If CAID3_Used = True And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & "," & fCAID3t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            headerStr = fCAID1t & headerStr
        End If
        
        If subprojectIDs Then 'v3.4.3
            headerStr = "Project," & headerStr
        End If
        
        Print #1, headerStr

        X = 1
        ActFound = False

        If curProj.Subprojects.Count > 0 Then

            Set subProjs = curProj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.Active = True And t.ExternalTask = False Then

                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If subprojectIDs Then 'v3.4.3
                                    ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                                End If
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
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
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
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).ID = ID
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
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

                    If t.Active = True And t.ExternalTask = False Then

                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If subprojectIDs Then 'v3.4.3
                                ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                            End If
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
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).CAM = CAM
                                ACTarray(X).ID = ID
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
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
                            ReDim Preserve ACTarray(1 To X)
                            If t.BaselineFinish <> "NA" Then
                                ACTarray(X).BFinish = t.BaselineFinish
                            End If
                            If t.BaselineStart <> "NA" Then
                                ACTarray(X).BStart = t.BaselineStart
                            End If
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).ID = ID
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            If subprojectIDs Then 'v3.4.3
                                ACTarray(X).SubProject = ProjID
                            End If
                            ACTarray(X).WP = WP
                            If t.ActualStart <> "NA" Then
                                ACTarray(X).AStart = t.ActualStart
                            End If
                            If t.ActualFinish <> "NA" Then
                                ACTarray(X).AFinish = t.ActualFinish
                            End If

                            X = X + 1
                            ActFound = True

nrETC_WP_Match_B:

                        End If

                    End If

                End If

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, dateFmt)
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, dateFmt)

                'v3.4.3 - refactored data output code
                
                outputStr = ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, dateFmt) & "," & Format(ACTarray(i).FFinish, dateFmt) & ","

                If aFinishString = "NA" Then
                    If CAID3_Used = True And CAID2_Used = True Then
                        outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAID3 & "," & outputStr
                    End If
                    If CAID3_Used = False And CAID2_Used = True Then
                        outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & outputStr
                    End If
                    If CAID3_Used = False And CAID2_Used = False Then
                        outputStr = ACTarray(i).CAID1 & "," & outputStr
                    End If
                    
                    If subprojectIDs Then 'v3.4.3
                        outputStr = ACTarray(i).SubProject & "," & outputStr
                    End If
                    
                    Print #1, outputStr
                End If

            Next i
        End If

        Close #1

    Else '**Resource Loaded**

        ACTfilename = destFolder & "\ETC ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
        RESfilename = destFolder & "\ETC RES_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        Open RESfilename For Output As #2

        'v3.4.3 - refactored header output code
        headerStr = ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date,"
        
        If CAID3_Used = True And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & "," & fCAID3t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            headerStr = fCAID1t & headerStr
        End If
        
        If subprojectIDs Then 'v3.4.3
            headerStr = "Project," & headerStr
        End If
        
        Print #1, headerStr

        Print #2, "Cobra ID,Resource,Amount,From Date,To Date"

        X = 1
        ActFound = False

        If curProj.Subprojects.Count > 0 Then

            Set subProjs = curProj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.Active = True And t.ExternalTask = False Then

                            If t.Work > 0 Or t.Cost > 0 Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If subprojectIDs Then 'v3.4.3
                                    ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                                End If
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
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If
                                    'v3.3.5 - check for ID length limit
                                    If Len(ID) > 58 Then
                                        ActIDCounter = ActIDCounter + 1
                                        ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                        ShortID = ACTarray(X).ShortID
                                    Else
                                        ACTarray(X).ShortID = ACTarray(X).ID
                                        ShortID = ACTarray(X).ShortID 'v3.3.6
                                    End If

                                    X = X + 1
                                    ActFound = True

                                    GoTo ETC_WP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        ShortID = ACTarray(i).ShortID
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
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).ID = ID
                                
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If
                                'v3.3.5 - check for ID length limit
                                If Len(ID) > 58 Then
                                    ActIDCounter = ActIDCounter + 1
                                    ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                    ShortID = ACTarray(X).ShortID
                                Else
                                    ACTarray(X).ShortID = ACTarray(X).ID
                                    ShortID = ACTarray(X).ShortID 'v3.3.6
                                End If

                                X = X + 1
                                ActFound = True

                                'Milestone Data
ETC_WP_Match:


                                Set tAss = t.Assignments

                                For Each tAssign In tAss

                                    If TimeScaleExport = True Then

                                        ExportTimeScaleResources ShortID, t, tAssign, 2, "ETC"

                                    Else

                                        Select Case tAssign.ResourceType

                                            Case pjResourceTypeWork

                                            If CStr(AssignmentResumeDate(tAssign)) <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingWork <> 0 Then 'v3.4.5

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork / 60 & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tAssign.Finish, dateFmt) 'v3.4.5

                                            ElseIf CStr(AssignmentResumeDate(tAssign)) = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingWork <> 0 Then 'v3.4.5

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                            End If

                                        Case pjResourceTypeCost

                                            If CStr(AssignmentResumeDate(tAssign)) <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingCost <> 0 Then 'v3.4.5

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingCost & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tAssign.Finish, dateFmt) 'v3.4.5
                                                

                                            ElseIf CStr(AssignmentResumeDate(tAssign)) = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingCost <> 0 Then 'v3.4.5

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                            End If

                                        Case pjResourceTypeMaterial

                                            If CStr(AssignmentResumeDate(tAssign)) <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingWork <> 0 Then 'v3.4.5

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tAssign.Finish, dateFmt) 'v3.4.5

                                            ElseIf CStr(AssignmentResumeDate(tAssign)) = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingWork <> 0 Then 'v3.4.5

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

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

                    If t.Active = True And t.ExternalTask = False Then

                        If t.Work > 0 Or t.Cost > 0 Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If subprojectIDs Then 'v3.4.3
                                ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                            End If
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
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).CAM = CAM
                                ACTarray(X).ID = ID
                                
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If
                                'v3.3.5 - check for ID length limit
                                If Len(ID) > 58 Then
                                    ActIDCounter = ActIDCounter + 1
                                    ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                    ShortID = ACTarray(X).ShortID
                                Else
                                    ACTarray(X).ShortID = ACTarray(X).ID
                                    ShortID = ACTarray(X).ShortID 'v3.3.6
                                End If

                                X = X + 1
                                ActFound = True

                                GoTo ETC_WP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    ShortID = ACTarray(i).ShortID
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
                            ReDim Preserve ACTarray(1 To X)
                            If t.BaselineFinish <> "NA" Then
                                ACTarray(X).BFinish = t.BaselineFinish
                            End If
                            If t.BaselineStart <> "NA" Then
                                ACTarray(X).BStart = t.BaselineStart
                            End If
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).ID = ID
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            If subprojectIDs Then 'v3.4.3
                                ACTarray(X).SubProject = ProjID
                            End If
                            ACTarray(X).WP = WP
                            If t.ActualStart <> "NA" Then
                                ACTarray(X).AStart = t.ActualStart
                            End If
                            If t.ActualFinish <> "NA" Then
                                ACTarray(X).AFinish = t.ActualFinish
                            End If
                            'v3.3.5 - check for ID length limit
                            If Len(ID) > 58 Then
                                ActIDCounter = ActIDCounter + 1
                                ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                ShortID = ACTarray(X).ShortID
                            Else
                                ACTarray(X).ShortID = ACTarray(X).ID
                                ShortID = ACTarray(X).ShortID 'v3.3.6
                            End If


                            X = X + 1
                            ActFound = True

                            'Milestone Data
ETC_WP_Match_B:


                            Set tAss = t.Assignments

                            For Each tAssign In tAss

                                If TimeScaleExport = True Then

                                    ExportTimeScaleResources ShortID, t, tAssign, 2, "ETC"
                                    
                                Else

                                    Select Case tAssign.ResourceType

                                        Case pjResourceTypeWork

                                        If CStr(AssignmentResumeDate(tAssign)) <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingWork <> 0 Then 'v3.4.5

                                            Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork / 60 & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tAssign.Finish, dateFmt) 'v3.4.5

                                        ElseIf CStr(AssignmentResumeDate(tAssign)) = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingWork <> 0 Then 'v3.4.5

                                            Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                        End If

                                    Case pjResourceTypeCost

                                        If CStr(AssignmentResumeDate(tAssign)) <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingCost <> 0 Then 'v3.4.5

                                            Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingCost & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tAssign.Finish, dateFmt) 'v3.4.5

                                        ElseIf CStr(AssignmentResumeDate(tAssign)) = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingCost <> 0 Then 'v3.4.5

                                            Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                        End If

                                    Case pjResourceTypeMaterial

                                        If CStr(AssignmentResumeDate(tAssign)) <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingWork <> 0 Then 'v3.4.5

                                            Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tAssign.Finish, dateFmt) 'v3.4.5

                                        ElseIf CStr(AssignmentResumeDate(tAssign)) = "NA" And tAssign.PercentWorkComplete <> 100 And tAssign.RemainingWork <> 0 Then 'v3.4.5

                                            Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

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

                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, dateFmt)
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, dateFmt)

                'v3.4.3 - refactored data output code
                
                outputStr = ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ShortID & "," & Format(ACTarray(i).FStart, dateFmt) & "," & Format(ACTarray(i).FFinish, dateFmt) & ","

                If aFinishString = "NA" Then
                    If CAID3_Used = True And CAID2_Used = True Then
                        outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAID3 & "," & outputStr
                    End If
                    If CAID3_Used = False And CAID2_Used = True Then
                        outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & outputStr
                    End If
                    If CAID3_Used = False And CAID2_Used = False Then
                        outputStr = ACTarray(i).CAID1 & "," & outputStr
                    End If
                    
                    If subprojectIDs Then 'v3.4.3
                        outputStr = ACTarray(i).SubProject & "," & outputStr
                    End If
                    
                    Print #1, outputStr
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
    Dim ProjID, CAID1, CAID3, WP, CAM, EVT, UID, CAID2, MSWeight, ID, ShortID, PCNT As String 'v3.3.5, v3.4.3
    Dim Milestone As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim WPDescArray() As WP_Descriptions
    Dim X As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String
    Dim headerStr As String 'v3.4.3
    Dim outputStr As String 'v3.4.3

    '*******************
    '****BCWS Export****
    '*******************
    
    ActIDCounter = 0 'v3.3.5

    If DescExport = True Then
        Get_WP_Descriptions curProj
    End If

    If ResourceLoaded = False Then

        ACTfilename = destFolder & "\BCWS ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        
        'v3.4.3 - refactored header output code
        headerStr = ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique,"
        
        If CAID3_Used = True And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & "," & fCAID3t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            headerStr = fCAID1t & headerStr
        End If
        
        If subprojectIDs Then 'v3.4.3
            headerStr = "Project," & headerStr
        End If
        
        Print #1, headerStr

        X = 1
        ActFound = False

        If curProj.Subprojects.Count > 0 Then

            Set subProjs = curProj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.Active = True And t.ExternalTask = False Then

                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If subprojectIDs Then 'v3.4.3
                                    ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                                End If
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
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

                                If Milestones_Used = False Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                    MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                End If

                                If BCRxport = True Then
                                    If IsInArray(WP, BCR_WP) = False Then
                                        GoTo Next_nrSProj_Task
                                    End If
                                End If

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    err.Raise 1
                                End If

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    ACTarray(X).CAID2 = CAID2
                                    ACTarray(X).WP = WP

                                    X = X + 1
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
                                ReDim Preserve ACTarray(1 To X)
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                ACTarray(X).CAID2 = CAID2
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                'Milestone Data
nrWP_Match:

                                If (EVT = "B" Or EVT = "B Milestone") And ExportMilestones Then
                                
                                    'v3.4.3 - refactored data output code
                                
                                    outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, dateFmt) & "," & Format(t.BaselineFinish, dateFmt) & ","

                                    If CAID3_Used = True And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        outputStr = CAID1 & "," & outputStr
                                    End If
                                    
                                    If subprojectIDs Then 'v3.4.3
                                        outputStr = ProjID & "," & outputStr
                                    End If
                                    
                                    Print #1, outputStr

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

                    If t.Active = True And t.ExternalTask = False Then

                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If subprojectIDs Then 'v3.4.3
                                ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                            End If
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))

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

                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            End If

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                err.Raise 1
                            End If

                            If BCRxport = True Then
                                If IsInArray(WP, BCR_WP) = False Then
                                    GoTo Next_nrTask
                                End If
                            End If

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                ACTarray(X).CAID2 = CAID2
                                ACTarray(X).WP = WP

                                X = X + 1
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
                            ReDim Preserve ACTarray(1 To X)
                            ACTarray(X).BFinish = t.BaselineFinish
                            ACTarray(X).BStart = t.BaselineStart
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            If subprojectIDs Then 'v3.4.3
                                ACTarray(X).SubProject = ProjID
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).ID = ID
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            ACTarray(X).CAID2 = CAID2
                            ACTarray(X).WP = WP

                            X = X + 1
                            ActFound = True

                            'Milestone Data
nrWP_Match_B:

                            If (EVT = "B" Or EVT = "B Milestone") And ExportMilestones Then
                            
                                'v3.4.3 - refactored data output code
                                
                                outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, dateFmt) & "," & Format(t.BaselineFinish, dateFmt) & ","
                            
                                If CAID3_Used = True And CAID2_Used = True Then
                                    outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    outputStr = CAID1 & "," & outputStr
                                End If
                                
                                If subprojectIDs Then 'v3.4.3
                                    outputStr = ProjID & "," & outputStr
                                End If
                                
                                Print #1, outputStr
                            End If

                        End If

                    End If

                End If
Next_nrTask:

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If DescExport = True Then
                    ACTarray(i).Desc = WP_Desc(ACTarray(i).ID)
                End If
                
                'v3.4.3 - refactored data output code
                
                outputStr = ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, dateFmt) & "," & Format(ACTarray(i).BFinish, dateFmt) & "," & ACTarray(i).EVT & ","

                If CAID3_Used = True And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAID3 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    outputStr = ACTarray(i).CAID1 & "," & outputStr
                End If
                
                If subprojectIDs Then 'v3.4.3
                    outputStr = ACTarray(i).SubProject & "," & outputStr
                End If
                
                Print #1, outputStr

            Next i
        End If

        Close #1

    Else

        ACTfilename = destFolder & "\BCWS ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
        RESfilename = destFolder & "\BCWS RES_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        Open RESfilename For Output As #2
        
        'v3.4.3 - refactored header output code
        headerStr = ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique,"
        
        If CAID3_Used = True And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & "," & fCAID3t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            headerStr = fCAID1t & headerStr
        End If
        
        If subprojectIDs Then 'v3.4.3
            headerStr = "Project," & headerStr
        End If
        
        Print #1, headerStr

        Print #2, "Cobra ID,Resource,Amount,From Date,To Date"

        X = 1
        ActFound = False

        If curProj.Subprojects.Count > 0 Then

            Set subProjs = curProj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.Active = True And t.ExternalTask = False Then

                            If t.BaselineWork > 0 Or t.BaselineCost > 0 Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If subprojectIDs Then 'v3.4.3
                                    ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                                End If
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))

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

                                If Milestones_Used = True Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                    MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                End If

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    err.Raise 1
                                End If

                                If BCRxport = True Then
                                    If IsInArray(WP, BCR_WP) = False Then
                                        GoTo Next_SProj_Task
                                    End If
                                End If

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    'v3.3.5 - check for ID length limit
                                    If Len(ID) > 58 Then
                                        ActIDCounter = ActIDCounter + 1
                                        ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                        ShortID = ACTarray(X).ShortID
                                    Else
                                        ACTarray(X).ShortID = ACTarray(X).ID
                                        ShortID = ACTarray(X).ShortID 'v3.3.6
                                    End If

                                    X = X + 1
                                    ActFound = True

                                    GoTo WP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        ShortID = ACTarray(i).ShortID
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
                                ReDim Preserve ACTarray(1 To X)
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                'v3.3.5 - check for ID length limit
                                If Len(ID) > 58 Then
                                    ActIDCounter = ActIDCounter + 1
                                    ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                    ShortID = ACTarray(X).ShortID
                                Else
                                    ACTarray(X).ShortID = ACTarray(X).ID
                                    ShortID = ACTarray(X).ShortID 'v3.3.6
                                End If

                                X = X + 1
                                ActFound = True

                                'Milestone Data
WP_Match:

                                If (EVT = "B" Or EVT = "B Milestone") And ExportMilestones Then
                                
                                    'v3.4.3 - refactored data output code
                                    
                                    outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, dateFmt) & "," & Format(t.BaselineFinish, dateFmt) & ","

                                    If CAID3_Used = True And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        outputStr = CAID1 & "," & outputStr
                                    End If
                                    
                                    If subprojectIDs Then 'v3.4.3
                                        outputStr = ProjID & "," & outputStr
                                    End If
                                    
                                    Print #1, outputStr

                                End If

                                Set tAss = t.Assignments

                                For Each tAssign In tAss

                                    If tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0 Then

                                        If TimeScaleExport = True Then

                                            ExportTimeScaleResources ShortID, t, tAssign, 2, "BCWS"

                                        Else

                                            Select Case tAssign.ResourceType

                                                Case pjResourceTypeWork

                                                    Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                                Case pjResourceTypeCost

                                                    Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                                Case pjResourceTypeMaterial

                                                    Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

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

                    If t.Active = True And t.ExternalTask = False Then

                        If t.BaselineWork > 0 Or t.BaselineCost > 0 Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If subprojectIDs Then 'v3.4.3
                                ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                            End If
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))

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

                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            End If

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                err.Raise 1
                            End If

                            If BCRxport = True Then
                                If IsInArray(WP, BCR_WP) = False Then
                                    GoTo next_task
                                End If
                            End If

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                'v3.3.5 - check for ID length limit
                                If Len(ID) > 58 Then
                                    ActIDCounter = ActIDCounter + 1
                                    ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                    ShortID = ACTarray(X).ShortID
                                Else
                                    ACTarray(X).ShortID = ACTarray(X).ID
                                    ShortID = ACTarray(X).ShortID 'v3.3.6
                                End If

                                X = X + 1
                                ActFound = True

                                GoTo WP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    ShortID = ACTarray(i).ShortID
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
                            ReDim Preserve ACTarray(1 To X)
                            ACTarray(X).BFinish = t.BaselineFinish
                            ACTarray(X).BStart = t.BaselineStart
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).ID = ID
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            If subprojectIDs Then 'v3.4.3
                                ACTarray(X).SubProject = ProjID
                            End If
                            ACTarray(X).WP = WP
                            'v3.3.5 - check for ID length limit
                            If Len(ID) > 58 Then
                                ActIDCounter = ActIDCounter + 1
                                ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                ShortID = ACTarray(X).ShortID
                            Else
                                ACTarray(X).ShortID = ACTarray(X).ID
                                ShortID = ACTarray(X).ShortID 'v3.3.6
                            End If

                            X = X + 1
                            ActFound = True

                            'Milestone Data
WP_Match_B:

                            If (EVT = "B" Or EVT = "B Milestone") And ExportMilestones Then
                            
                                'v3.4.3 - refactored data output code
                                
                                outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, dateFmt) & "," & Format(t.BaselineFinish, dateFmt) & ","
                            
                                If CAID3_Used = True And CAID2_Used = True Then
                                    outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    outputStr = CAID1 & "," & outputStr
                                End If
                                
                                If subprojectIDs Then 'v3.4.3
                                    outputStr = ProjID & "," & outputStr
                                End If
                                
                                Print #1, outputStr
                            End If

                            Set tAss = t.Assignments

                            For Each tAssign In tAss

                                If tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0 Then

                                    If TimeScaleExport = True Then

                                        ExportTimeScaleResources ShortID, t, tAssign, 2, "BCWS"

                                    Else

                                        Select Case tAssign.ResourceType

                                            Case pjResourceTypeWork

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                            Case pjResourceTypeCost

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                            Case pjResourceTypeMaterial

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

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

                If DescExport = True Then
                    ACTarray(i).Desc = WP_Desc(ACTarray(i).ID)
                End If
                
                'v3.4.3 - refactored data output code
                
                outputStr = ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ShortID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, dateFmt) & "," & Format(ACTarray(i).BFinish, dateFmt) & "," & ACTarray(i).EVT & ","

                If CAID3_Used = True And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAID3 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    outputStr = ACTarray(i).CAID1 & "," & outputStr
                End If
                
                If subprojectIDs Then 'v3.4.3
                    outputStr = ACTarray(i).SubProject & "," & outputStr
                End If
                
                Print #1, outputStr

            Next i
        End If

        Close #1
        Close #2

    End If
        
End Sub

Private Sub WhatIf_Export(ByVal curProj As Project) 'v3.2

    Dim t As Task
    Dim tAss As Assignments
    Dim tAssign As Assignment
    Dim ProjID, CAID1, CAID3, WP, CAM, EVT, UID, CAID2, MSWeight, ID, ShortID, PCNT As String 'v3.3.5, v3.4.3
    Dim Milestone As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim WPDescArray() As WP_Descriptions
    Dim X As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String
    Dim headerStr As String 'v3.4.3
    Dim outputStr As String 'v3.4.3

    '*******************
    '**What-if Export***
    '*******************
    
    ActIDCounter = 0 'v3.3.5

    If DescExport = True Then
        Get_WP_Descriptions curProj
    End If

    If ResourceLoaded = False Then

        ACTfilename = destFolder & "\WhatIf ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        
        'v3.4.3 - refactored header output code
        headerStr = ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique,"
        
        If CAID3_Used = True And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & "," & fCAID3t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            headerStr = fCAID1t & headerStr
        End If
        
        If subprojectIDs Then 'v3.4.3
            headerStr = "Project," & headerStr
        End If
        
        Print #1, headerStr

        X = 1
        ActFound = False

        If curProj.Subprojects.Count > 0 Then

            Set subProjs = curProj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.Active = True And t.ExternalTask = False Then

                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "D" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "d" Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If subprojectIDs Then 'v3.4.3
                                    ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                                End If
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
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

                                If Milestones_Used = False Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                    MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                End If

                                If BCRxport = True Then
                                    If IsInArray(WP, BCR_WP) = False Then
                                        GoTo Next_nrSProj_Task
                                    End If
                                End If

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    err.Raise 1
                                End If

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                        ACTarray(X).BStart = t.BaselineStart
                                    Else
                                        ACTarray(X).BFinish = t.Finish
                                        ACTarray(X).BStart = t.Start
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    ACTarray(X).CAID2 = CAID2
                                    ACTarray(X).WP = WP

                                    X = X + 1
                                    ActFound = True

                                    GoTo nrWP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        
                                        'Found an existing matching WP line
                                        If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
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
                                        Else
                                            If ACTarray(i).BStart > t.Start Then
                                                ACTarray(i).BStart = t.Start
                                            End If
                                            If ACTarray(i).BFinish < t.Finish Then
                                                ACTarray(i).BFinish = t.Finish
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                        End If
                                        GoTo nrWP_Match
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                Else
                                    ACTarray(X).BFinish = t.Finish
                                    ACTarray(X).BStart = t.Start
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                ACTarray(X).CAID2 = CAID2
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                'Milestone Data
nrWP_Match:

                                If (EVT = "B" Or EVT = "B Milestone") And ExportMilestones Then
                                    'v3.4.3 - refactored data output code
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    
                                        outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, dateFmt) & "," & Format(t.BaselineFinish, dateFmt) & ","
                                    
                                        If CAID3_Used = True And CAID2_Used = True Then
                                            outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                        End If
                                        If CAID3_Used = False And CAID2_Used = True Then
                                            outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                        End If
                                        If CAID3_Used = False And CAID2_Used = False Then
                                            outputStr = CAID1 & "," & outputStr
                                        End If
                                        
                                        If subprojectIDs Then 'v3.4.3
                                            outputStr = ProjID & "," & outputStr
                                        End If
                                        
                                        Print #1, outputStr
                                    Else
                                    
                                        outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, dateFmt) & "," & Format(t.Finish, dateFmt) & ","
                                    
                                        If CAID3_Used = True And CAID2_Used = True Then
                                            outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                        End If
                                        If CAID3_Used = False And CAID2_Used = True Then
                                            outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                        End If
                                        If CAID3_Used = False And CAID2_Used = False Then
                                            outputStr = CAID1 & "," & outputStr
                                        End If
                                        
                                        If subprojectIDs Then 'v3.4.3
                                            outputStr = ProjID & "," & outputStr
                                        End If
                                        
                                        Print #1, outputStr
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

                    If t.Active = True And t.ExternalTask = False Then
                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "D" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "d" Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If subprojectIDs Then 'v3.4.3
                                ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                            End If
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))

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

                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            End If

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                err.Raise 1
                            End If

                            If BCRxport = True Then
                                If IsInArray(WP, BCR_WP) = False Then
                                    GoTo Next_nrTask
                                End If
                            End If

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                Else
                                    ACTarray(X).BFinish = t.Finish
                                    ACTarray(X).BStart = t.Start
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                ACTarray(X).CAID2 = CAID2
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                GoTo nrWP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
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
                                    Else
                                        If ACTarray(i).BStart > t.Start Then
                                            ACTarray(i).BStart = t.Start
                                        End If
                                        If ACTarray(i).BFinish < t.Finish Then
                                            ACTarray(i).BFinish = t.Finish
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                    End If
                                    GoTo nrWP_Match_B
                                End If
                            Next i

                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To X)
                            If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                            Else
                                ACTarray(X).BFinish = t.Finish
                                ACTarray(X).BStart = t.Start
                            End If
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            If subprojectIDs Then 'v3.4.3
                                ACTarray(X).SubProject = ProjID
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).ID = ID
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            ACTarray(X).CAID2 = CAID2
                            ACTarray(X).WP = WP

                            X = X + 1
                            ActFound = True

                            'Milestone Data
nrWP_Match_B:

                            If (EVT = "B" Or EVT = "B Milestone") And ExportMilestones Then
                                'v3.4.3 - refactored data output code
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                
                                    outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, dateFmt) & "," & Format(t.BaselineFinish, dateFmt) & ","
                                
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        outputStr = CAID1 & "," & outputStr
                                    End If
                                    
                                    If subprojectIDs Then 'v3.4.3
                                        outputStr = ProjID & "," & outputStr
                                    End If
                                    
                                    Print #1, outputStr
                                Else
                                    
                                    outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, dateFmt) & "," & Format(t.Finish, dateFmt) & ","
                                
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        outputStr = CAID1 & "," & outputStr
                                    End If
                                    
                                    If subprojectIDs Then 'v3.4.3
                                        outputStr = ProjID & "," & outputStr
                                    End If
                                    
                                    Print #1, outputStr
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

                If DescExport = True Then
                    ACTarray(i).Desc = WP_Desc(ACTarray(i).ID)
                End If

                'v3.4.3 - refactored data output code
                
                outputStr = ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, dateFmt) & "," & Format(ACTarray(i).BFinish, dateFmt) & "," & ACTarray(i).EVT & ","

                If CAID3_Used = True And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAID3 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    outputStr = ACTarray(i).CAID1 & "," & outputStr
                End If
                
                If subprojectIDs Then 'v3.4.3
                    outputStr = ACTarray(i).SubProject & "," & outputStr
                End If
                
                Print #1, outputStr

            Next i
        End If

        Close #1

    Else

        ACTfilename = destFolder & "\WhatIf ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
        RESfilename = destFolder & "\WhatIf RES_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Project) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        Open RESfilename For Output As #2
        
        'v3.4.3 - refactored header output code
        headerStr = ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique,"
        
        If CAID3_Used = True And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & "," & fCAID3t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            headerStr = fCAID1t & "," & fCAID2t & headerStr
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            headerStr = fCAID1t & headerStr
        End If
        
        If subprojectIDs Then 'v3.4.3
            headerStr = "Project," & headerStr
        End If
        
        Print #1, headerStr

        Print #2, "Cobra ID,Resource,Amount,From Date,To Date"

        X = 1
        ActFound = False

        If curProj.Subprojects.Count > 0 Then

            Set subProjs = curProj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.Active = True And t.ExternalTask = False Then
                        
                            If ((t.BaselineWork > 0 Or t.BaselineCost > 0) And _
                            (t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "d" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "D")) _
                            Or _
                            ((t.Work > 0 Or t.Cost > 0) And _
                            (t.GetField(FieldNameToFieldConstant(fWhatIf)) = "R" Or t.GetField(FieldNameToFieldConstant(fWhatIf)) = "r")) Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If subprojectIDs Then 'v3.4.3
                                    ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                                End If
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))

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

                                If Milestones_Used = True Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                    MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                End If

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    err.Raise 1
                                End If

                                If BCRxport = True Then
                                    If IsInArray(WP, BCR_WP) = False Then
                                        GoTo Next_SProj_Task
                                    End If
                                End If

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                        ACTarray(X).BStart = t.BaselineStart
                                    Else
                                        ACTarray(X).BFinish = t.Finish
                                        ACTarray(X).BStart = t.Start
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    If subprojectIDs Then 'v3.4.3
                                        ACTarray(X).SubProject = ProjID
                                    End If
                                    ACTarray(X).WP = WP
                                    'v3.3.5 - check for ID length limit
                                    If Len(ID) > 58 Then
                                        ActIDCounter = ActIDCounter + 1
                                        ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                        ShortID = ACTarray(X).ShortID
                                    Else
                                        ACTarray(X).ShortID = ACTarray(X).ID
                                        ShortID = ACTarray(X).ShortID 'v3.3.6
                                    End If

                                    X = X + 1
                                    ActFound = True

                                    GoTo WP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        ShortID = ACTarray(i).ShortID
                                        'Found an existing matching WP line
                                        If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                            If ACTarray(i).BStart > t.BaselineStart Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            End If
                                            If ACTarray(i).BFinish < t.BaselineFinish Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            End If
                                        Else
                                            If ACTarray(i).BStart > t.Start Then
                                                ACTarray(i).BStart = t.Start
                                            End If
                                            If ACTarray(i).BFinish < t.Finish Then
                                                ACTarray(i).BFinish = t.Finish
                                            End If

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
                                ReDim Preserve ACTarray(1 To X)
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                Else
                                    ACTarray(X).BFinish = t.Finish
                                    ACTarray(X).BStart = t.Start
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                'v3.3.5 - check for ID length limit
                                If Len(ID) > 58 Then
                                    ActIDCounter = ActIDCounter + 1
                                    ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                    ShortID = ACTarray(X).ShortID
                                Else
                                    ACTarray(X).ShortID = ACTarray(X).ID
                                    ShortID = ACTarray(X).ShortID 'v3.3.6
                                End If

                                X = X + 1
                                ActFound = True

                                'Milestone Data
WP_Match:

                                If (EVT = "B" Or EVT = "B Milestone") And ExportMilestones Then
                                    'v3.4.3 - refactored data output code
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        
                                        outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, dateFmt) & "," & Format(t.BaselineFinish, dateFmt) & ","
                                        
                                        If CAID3_Used = True And CAID2_Used = True Then
                                            outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                        End If
                                        If CAID3_Used = False And CAID2_Used = True Then
                                            outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                        End If
                                        If CAID3_Used = False And CAID2_Used = False Then
                                            outputStr = CAID1 & "," & outputStr
                                        End If
                                        
                                        If subprojectIDs Then 'v3.4.3
                                            outputStr = ProjID & "," & outputStr
                                        End If
                                        
                                        Print #1, outputStr
                                        
                                    Else
                                        
                                        outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, dateFmt) & "," & Format(t.Finish, dateFmt) & ","
                                        
                                        If CAID3_Used = True And CAID2_Used = True Then
                                            outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                        End If
                                        If CAID3_Used = False And CAID2_Used = True Then
                                            outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                        End If
                                        If CAID3_Used = False And CAID2_Used = False Then
                                            outputStr = CAID1 & "," & outputStr
                                        End If
                                        
                                        If subprojectIDs Then 'v3.4.3
                                            outputStr = ProjID & "," & outputStr
                                        End If
                                        
                                        Print #1, outputStr
                                        
                                    End If
                                End If

                                Set tAss = t.Assignments

                                For Each tAssign In tAss

                                    If (tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0) And _
                                    (t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R") Then

                                        If TimeScaleExport = True Then

                                            ExportTimeScaleResources ShortID, t, tAssign, 2, "BCWS"

                                        Else

                                            Select Case tAssign.ResourceType

                                                Case pjResourceTypeWork

                                                    Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                                Case pjResourceTypeCost

                                                    Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                                Case pjResourceTypeMaterial

                                                    Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                            End Select

                                        End If
                                    Else
                                    
                                        If (tAssign.Work <> 0 Or tAssign.Cost <> 0) And _
                                        (t.GetField(FieldNameToFieldConstant(fWhatIf)) = "r" Or t.GetField(FieldNameToFieldConstant(fWhatIf)) = "R") Then

                                            If TimeScaleExport = True Then
    
                                                ExportTimeScaleResources ShortID, t, tAssign, 2, "ETC"
    
                                            Else
    
                                                Select Case tAssign.ResourceType
    
                                                    Case pjResourceTypeWork
    
                                                        Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)
    
                                                    Case pjResourceTypeCost
    
                                                        Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)
    
                                                    Case pjResourceTypeMaterial
    
                                                        Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)
    
                                                End Select
    
                                            End If
                                            
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

                    If t.Active = True And t.ExternalTask = False Then

                        If ((t.BaselineWork > 0 Or t.BaselineCost > 0) And _
                        (t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "d" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "D")) _
                        Or _
                        ((t.Work > 0 Or t.Cost > 0) And _
                        (t.GetField(FieldNameToFieldConstant(fWhatIf)) = "R" Or t.GetField(FieldNameToFieldConstant(fWhatIf)) = "r")) Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If subprojectIDs Then 'v3.4.3
                                ProjID = t.GetField(FieldNameToFieldConstant(fProject))
                            End If
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))

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

                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            End If

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                err.Raise 1
                            End If

                            If BCRxport = True Then
                                If IsInArray(WP, BCR_WP) = False Then
                                    GoTo next_task
                                End If
                            End If

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                Else
                                    ACTarray(X).BFinish = t.Finish
                                    ACTarray(X).BStart = t.Start
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                If subprojectIDs Then 'v3.4.3
                                    ACTarray(X).SubProject = ProjID
                                End If
                                ACTarray(X).WP = WP
                                'v3.3.5 - check for ID length limit
                                If Len(ID) > 58 Then
                                    ActIDCounter = ActIDCounter + 1
                                    ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                    ShortID = ACTarray(X).ShortID
                                Else
                                    ACTarray(X).ShortID = ACTarray(X).ID
                                    ShortID = ACTarray(X).ShortID 'v3.3.6
                                End If

                                X = X + 1
                                ActFound = True

                                GoTo WP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    ShortID = ACTarray(i).ShortID
                                    'Found an existing matching WP line
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        If ACTarray(i).BStart > t.BaselineStart Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        End If
                                        If ACTarray(i).BFinish < t.BaselineFinish Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        End If
                                    Else
                                        If ACTarray(i).BStart > t.Start Then
                                            ACTarray(i).BStart = t.Start
                                        End If
                                        If ACTarray(i).BFinish < t.Finish Then
                                            ACTarray(i).BFinish = t.Finish
                                        End If

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
                            ReDim Preserve ACTarray(1 To X)
                            If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                            Else
                                ACTarray(X).BFinish = t.Finish
                                ACTarray(X).BStart = t.Start
                            End If
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).ID = ID
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            If subprojectIDs Then 'v3.4.3
                                ACTarray(X).SubProject = ProjID
                            End If
                            ACTarray(X).WP = WP
                            'v3.3.5 - check for ID length limit
                            If Len(ID) > 58 Then
                                ActIDCounter = ActIDCounter + 1
                                ACTarray(X).ShortID = ACTarray(X).WP '& " (" & ActIDCounter & ")"
                                ShortID = ACTarray(X).ShortID
                            Else
                                ACTarray(X).ShortID = ACTarray(X).ID
                                ShortID = ACTarray(X).ShortID 'v3.3.6
                            End If

                            X = X + 1
                            ActFound = True

                            'Milestone Data
WP_Match_B:

                            If (EVT = "B" Or EVT = "B Milestone") And ExportMilestones Then
                                'v3.4.3 - refactored data output code
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        
                                    outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, dateFmt) & "," & Format(t.BaselineFinish, dateFmt) & ","
                                        
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        outputStr = CAID1 & "," & outputStr
                                    End If
                                    
                                    If subprojectIDs Then 'v3.4.3
                                        outputStr = ProjID & "," & outputStr
                                    End If
                                    
                                    Print #1, outputStr
                                    
                                Else
                                    
                                    outputStr = CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, dateFmt) & "," & Format(t.Finish, dateFmt) & ","
                                    
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & CAID3 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        outputStr = CAID1 & "," & CAID2 & "," & outputStr
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        outputStr = CAID1 & "," & outputStr
                                    End If
                                    
                                    If subprojectIDs Then 'v3.4.3
                                        outputStr = ProjID & "," & outputStr
                                    End If
                                    
                                    Print #1, outputStr
                                    
                                End If
                            End If

                            Set tAss = t.Assignments

                            For Each tAssign In tAss

                                If (tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0) And _
                                (t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R") Then

                                    If TimeScaleExport = True Then

                                        ExportTimeScaleResources ShortID, t, tAssign, 2, "BCWS"

                                    Else

                                        Select Case tAssign.ResourceType

                                            Case pjResourceTypeWork

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                            Case pjResourceTypeCost

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                            Case pjResourceTypeMaterial

                                                Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                        End Select

                                    End If
                                    
                                Else
                                    
                                    If (tAssign.Work <> 0 Or tAssign.Cost <> 0) And _
                                    (t.GetField(FieldNameToFieldConstant(fWhatIf)) = "r" Or t.GetField(FieldNameToFieldConstant(fWhatIf)) = "R") Then

                                        If TimeScaleExport = True Then
    
                                            ExportTimeScaleResources ShortID, t, tAssign, 2, "ETC"
    
                                        Else
    
                                            Select Case tAssign.ResourceType
    
                                                Case pjResourceTypeWork
    
                                                    Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)
    
                                                Case pjResourceTypeCost
    
                                                    Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)
    
                                                Case pjResourceTypeMaterial
    
                                                    Print #2, ShortID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)
    
                                            End Select
    
                                        End If
                                    
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

                If DescExport = True Then
                    ACTarray(i).Desc = WP_Desc(ACTarray(i).ID)
                End If

                'v3.4.3 - refactored data output code
                
                outputStr = ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ShortID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, dateFmt) & "," & Format(ACTarray(i).BFinish, dateFmt) & "," & ACTarray(i).EVT & ","

                If CAID3_Used = True And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAID3 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    outputStr = ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & outputStr
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    outputStr = ACTarray(i).CAID1 & "," & outputStr
                End If
                
                If subprojectIDs Then 'v3.4.3
                    outputStr = ACTarray(i).SubProject & "," & outputStr
                End If
                
                Print #1, outputStr

            Next i
        End If

        Close #1
        Close #2

    End If
        
End Sub

Private Function SetDirectory(ByVal ProjName As String) As String
    Dim newDir As String
    Dim pathDesktop As String

    pathDesktop = BrowseForFolder 'CreateObject("WScript.Shell").SpecialFolders("Desktop")) 'v3.4.2
    
    newDir = pathDesktop & "\" & RemoveIllegalCharacters(ProjName) & "_" & Format(Now, "YYYYMMDD HHMMSS")

    If Len(newDir) > 220 Then 'v3.4.1
        newDir = "\\?\" & newDir
    End If

    MkDir newDir
    SetDirectory = CreateObject("Scripting.FileSystemObject").GetFolder(newDir).ShortPath 'v3.4.1
    Exit Function

End Function

Private Sub DeleteDirectory(ByVal DirName As String)

    RmDir DirName & "\"
    Exit Sub

End Sub

Private Function WP_Desc(ByVal WP_ID As String) As String

    Dim tempDesc As String
    Dim X As Integer

    tempDesc = ""

    If WPDescCount = 0 Then
        WP_Desc = tempDesc
        Exit Function
    End If

    On Error GoTo NoWPMatchFound

    For X = 1 To UBound(WPDescArray)
        If WPDescArray(X).WP_ID = WP_ID Then
            tempDesc = WPDescArray(X).Desc
            WP_Desc = tempDesc
            Exit Function
        End If
    Next X

NoWPMatchFound:

    WP_Desc = ""

End Function

Private Sub Get_WP_Descriptions(ByVal curProj As Project)

    Dim CAID1 As String
    Dim CAID2 As String
    Dim CAID3 As String
    Dim WP As String
    Dim ID As String
    Dim Desc As String
    Dim i As Integer
    Dim X As Integer
    '<issue47>
    Dim subProjs As Subprojects
    Dim subProj As SubProject
    Dim curSProj As Project
    Dim t As Task '</issue47>

    WPDescCount = 0

    i = 0

    If curProj.Subprojects.Count > 0 Then

        Set subProjs = curProj.Subprojects

        For Each subProj In subProjs

            FileOpen Name:=subProj.Path, ReadOnly:=True

            Set curSProj = ActiveProject

            For Each t In curSProj.Tasks '<issue47>

                If Not t Is Nothing Then

                    WP = t.GetField(FieldNameToFieldConstant(fWP))

                    If WP = "" Then GoTo Next_SubProj_WPtask

                    CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                    If CAID3_Used = True Then
                        CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                    End If
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
                    Desc = Replace(t.Name, ",", "")

                    If i = 0 Then
                        i = 1
                    End If

                    If i = 1 Then

                        ReDim WPDescArray(1 To i)
                        WPDescArray(i).WP_ID = ID
                        WPDescArray(i).Desc = Desc
                        WPDescCount = i
                        i = i + 1

                    Else

                        For X = 1 To UBound(WPDescArray)

                            If WPDescArray(X).WP_ID = ID Then
                                GoTo Next_SubProj_WPtask
                            End If
                        Next X

                        ReDim Preserve WPDescArray(1 To i)
                        WPDescArray(i).WP_ID = ID
                        WPDescArray(i).Desc = Desc
                        WPDescCount = i
                        i = i + 1

                    End If

                End If
Next_SubProj_WPtask:

            Next t

        Next subProj

    Else

        For Each t In curProj.Tasks

            If Not t Is Nothing Then
                WP = t.GetField(FieldNameToFieldConstant(fWP))

                If WP = "" Then GoTo Next_WPtask

                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                If CAID3_Used = True Then
                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                End If
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
                Desc = Replace(t.Name, ",", "")

                If i = 0 Then
                    i = 1
                End If

                If i = 1 Then

                    ReDim WPDescArray(1 To i)
                    WPDescArray(i).WP_ID = ID
                    WPDescArray(i).Desc = Desc
                    WPDescCount = i
                    i = i + 1

                Else

                    For X = 1 To UBound(WPDescArray)

                        If WPDescArray(X).WP_ID = ID Then
                            GoTo Next_WPtask
                        End If
                    Next X

                    ReDim Preserve WPDescArray(1 To i)
                    WPDescArray(i).WP_ID = ID
                    WPDescArray(i).Desc = Desc
                    WPDescCount = i
                    i = i + 1

                End If

            End If

Next_WPtask:

        Next t

    End If

End Sub

Private Function IsInArray(ByVal stringToBeFound As String, ByVal arr As Variant) As Boolean
'v3.3.11 - rewrote function to mitigate false positives on null WP strings

    Dim i As Integer
    
    For i = 1 To UBound(arr)
        
        If arr(i) = stringToBeFound Then
        
            IsInArray = True
            Exit Function
        
        End If
        
    Next i
    
    IsInArray = False
End Function

Private Function CleanCamName(ByVal CAM As String) As String

    Dim tempCAM As String

    tempCAM = CAM

    If InStr(tempCAM, ".") > 0 Then
        tempCAM = Right(tempCAM, Len(tempCAM) - InStr(tempCAM, "."))
    End If

    If InStr(tempCAM, ",") > 0 Then
        tempCAM = Replace(tempCAM, ",", " ")
    End If

    CleanCamName = tempCAM

End Function

Private Function Find_BCRs(ByVal curProj As Project, ByVal fWP As String, ByVal fBCR As String, ByVal BCRnum As String) As Integer

    Dim t As Task
    Dim i As Integer
    Dim X As Integer
    Dim tempBCRstr As String
    Dim subProjs As Subprojects
    Dim subProj As SubProject
    Dim curSProj As Project

    i = 0

    If curProj.Subprojects.Count > 0 Then

        Set subProjs = curProj.Subprojects

        For Each subProj In subProjs

            FileOpen Name:=subProj.Path, ReadOnly:=True

            Set curSProj = ActiveProject

            For Each t In curSProj.Tasks

                If Not t Is Nothing Then
                
                    If t.Active = True Then 'todo: should this have ExternalTask = False also?
                    '3.3.12: ignore summary and inactive tasks
                    
                        tempBCRstr = t.GetField(FieldNameToFieldConstant(fBCR))
    
                        If InStr(tempBCRstr, BCRnum) > 0 Then
    
                            If i = 0 Then
                                i = 1
                            End If
    
                            If i = 1 Then
    
                                ReDim BCR_WP(1 To i)
                                BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                                Debug.Print "Storing BCR ID: " & BCR_WP(i)
                                i = i + 1
    
                            Else
    
                                For X = 1 To UBound(BCR_WP)
                                    If BCR_WP(X) = t.GetField(FieldNameToFieldConstant(fWP)) Then
                                        GoTo Next_SubProj_WPtask
                                    End If
                                Next X
    
                                ReDim Preserve BCR_WP(1 To i)
                                BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                                Debug.Print "Storing BCR ID: " & BCR_WP(i)
                                i = i + 1
    
                            End If
    
                        End If
                        
                    End If
                    
                End If
Next_SubProj_WPtask:

            Next t

        Next subProj

    Else

        For Each t In curProj.Tasks

            If Not t Is Nothing Then
            
                If t.Active = True Then 'todo: should this have ExternalTask = False also?
                '3.3.12: ignore summary and inactive tasks
                
                    tempBCRstr = t.GetField(FieldNameToFieldConstant(fBCR))
    
                    If InStr(tempBCRstr, BCRnum) > 0 Then
    
                        If i = 0 Then
                            i = 1
                        End If
    
                        If i = 1 Then
    
                            ReDim BCR_WP(1 To i)
                            BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                            Debug.Print "Storing BCR ID: " & BCR_WP(i)
                            i = i + 1
    
                        Else
    
                            For X = 1 To UBound(BCR_WP)
                                If BCR_WP(X) = t.GetField(FieldNameToFieldConstant(fWP)) Then
                                    GoTo Next_WPtask
                                End If
                            Next X
    
                            ReDim Preserve BCR_WP(1 To i)
                            BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                            Debug.Print "Storing BCR ID: " & BCR_WP(i)
                            i = i + 1
    
                        End If
    
                    End If
                    
                End If

            End If

Next_WPtask:

        Next t

    End If

    Find_BCRs = i

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

    Dim tsv As TimeScaleValue
    Dim tsvs As TimeScaleValues
    Dim tsvsa As TimeScaleValues
    Dim tsvA As TimeScaleValue
    Dim tempWork As Double

    Select Case exportType

        Case "ETC"

            Select Case tAssign.ResourceType

                Case pjResourceTypeWork

                    If CStr(AssignmentResumeDate(tAssign)) <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then 'v3.4.5

                        If TsvScale = "Weekly" Then 'v3.4
                            Set tsvs = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks) 'v3.4.5
                            Set tsvsa = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledActualWork, pjTimescaleWeeks) 'v3.4.5
                        Else
                            Set tsvs = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleMonths) 'v3.4, v3.4.5
                            Set tsvsa = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledActualWork, pjTimescaleMonths) 'v3.4, v3.4.5
                        End If
                        
                        For Each tsv In tsvs

                            Set tsvA = tsvsa(tsv.Index)

                            tempWork = 0

                            If tsvA <> "" Then
                                tempWork = CDbl(tsv.Value) - CDbl(tsvA.Value) 'v3.3.6
                            ElseIf tsv.Value <> "" Then
                                tempWork = CDbl(tsv.Value)
                            End If

                            If tempWork <> 0 Then

                                If tsvs.Count = 1 Then

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tAssign.Finish, dateFmt) 'v3.4.5

                                Else

                                    Select Case tsv.Index

                                        Case 1

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt) 'v3.4.5

                                        Case tsvs.Count

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                        Case Else

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                    End Select

                                End If

                            End If

                        Next tsv

                        Exit Sub

                    ElseIf CStr(AssignmentResumeDate(tAssign)) = "NA" And tAssign.PercentWorkComplete <> 100 Then 'v3.4.5

                        If TsvScale = "Weekly" Then 'v3.4
                            Set tsvs = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        Else
                            Set tsvs = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleMonths) 'v3.4
                        End If
                        
                        For Each tsv In tsvs

                            If tsv.Value <> "" Then

                                If tsvs.Count = 1 Then

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                Else

                                    Select Case tsv.Index

                                        Case 1

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tAssign.Start, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                        Case tsvs.Count

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                        Case Else

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                    End Select

                                End If

                            End If

                        Next tsv

                        Exit Sub

                    End If

            Case pjResourceTypeCost

                If CStr(AssignmentResumeDate(tAssign)) <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then 'v3.4.5

                    If TsvScale = "Weekly" Then 'v3.4
                        Set tsvs = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledCost, pjTimescaleWeeks) 'v3.4.5
                        Set tsvsa = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledActualCost, pjTimescaleWeeks) 'v3.4.5
                    Else
                        Set tsvs = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledCost, pjTimescaleMonths) 'v3.4, v3.4.5
                        Set tsvsa = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledActualCost, pjTimescaleMonths) 'v3.4, v3.4.5
                    End If
                    
                    For Each tsv In tsvs

                        Set tsvA = tsvsa(tsv.Index)

                        tempWork = 0

                        If tsvA <> "" Then
                            tempWork = CDbl(tsv.Value) - CDbl(tsvA.Value) 'v3.3.6
                        ElseIf tsv.Value <> "" Then
                            tempWork = CDbl(tsv.Value)
                        End If

                        If tempWork <> 0 Then

                            If tsvs.Count = 1 Then

                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tAssign.Finish, dateFmt) 'v3.4.5

                            Else

                                Select Case tsv.Index

                                    Case 1

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt) 'v3.4.5

                                    Case tsvs.Count

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                    Case Else

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                End Select

                            End If

                        End If

                    Next tsv

                    Exit Sub

                ElseIf CStr(AssignmentResumeDate(tAssign)) = "NA" And tAssign.PercentWorkComplete <> 100 Then 'v3.4.5

                    If TsvScale = "Weekly" Then 'v3.4
                        Set tsvs = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledCost, pjTimescaleWeeks)
                    Else
                        Set tsvs = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledCost, pjTimescaleMonths) 'v3.4
                    End If
                    
                    For Each tsv In tsvs

                        If tsv.Value <> "" Then

                            If tsvs.Count = 1 Then

                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                            Else

                                Select Case tsv.Index

                                    Case 1

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.Start, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                    Case tsvs.Count

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                    Case Else

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                End Select

                            End If

                        End If

                    Next tsv

                    Exit Sub

                End If

            Case pjResourceTypeMaterial

                If CStr(AssignmentResumeDate(tAssign)) <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then 'v3.4.5

                        If TsvScale = "Weekly" Then 'v3.4
                            Set tsvs = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks) 'v3.4.5
                            Set tsvsa = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledActualWork, pjTimescaleWeeks) 'v3.4.5
                        Else
                            Set tsvs = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleMonths) 'v3.4, v3.4.5
                            Set tsvsa = tAssign.TimeScaleData(AssignmentResumeDate(tAssign), tAssign.Finish, pjAssignmentTimescaledActualWork, pjTimescaleMonths) 'v3.4, v3.4.5
                        End If

                        For Each tsv In tsvs

                            Set tsvA = tsvsa(tsv.Index)

                            tempWork = 0

                            If tsvA <> "" Then
                                tempWork = CDbl(tsv.Value) - CDbl(tsvA.Value) 'v3.3.6
                            ElseIf tsv.Value <> "" Then
                                tempWork = CDbl(tsv.Value)
                            End If

                            If tempWork <> 0 Then

                                If tsvs.Count = 1 Then

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tAssign.Finish, dateFmt) 'v3.4.5

                                Else

                                    Select Case tsv.Index

                                        Case 1

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(AssignmentResumeDate(tAssign), dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt) 'v3.4.5

                                        Case tsvs.Count

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                        Case Else

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                    End Select

                                End If

                            End If

                        Next tsv

                        Exit Sub

                    ElseIf CStr(AssignmentResumeDate(tAssign)) = "NA" And tAssign.PercentWorkComplete <> 100 Then 'v3.4.5

                        If TsvScale = "Weekly" Then 'v3.4
                            Set tsvs = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        Else
                            Set tsvs = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleMonths) 'v3.4
                        End If
                        
                        For Each tsv In tsvs

                            If tsv.Value <> "" Then

                                If tsvs.Count = 1 Then

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.Start, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                Else

                                    Select Case tsv.Index

                                        Case 1

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.Start, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                        Case tsvs.Count

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tAssign.Finish, dateFmt)

                                        Case Else

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                    End Select

                                End If

                            End If

                        Next tsv

                        Exit Sub

                    End If

            End Select

        Case "BCWS"

            Select Case tAssign.ResourceType

                Case pjResourceTypeWork

                    If TsvScale = "Weekly" Then 'v3.4
                        Set tsvs = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                    Else
                        Set tsvs = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleMonths) 'v3.4
                    End If
                    
                    For Each tsv In tsvs

                        If tsv.Value <> "" Then

                            If tsvs.Count = 1 Then

                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                            Else

                                Select Case tsv.Index

                                    Case 1

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                    Case tsvs.Count

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                    Case Else

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                End Select

                            End If

                        End If

                    Next tsv

                    Exit Sub

            Case pjResourceTypeCost

                If TsvScale = "Weekly" Then 'v3.4
                    Set tsvs = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineCost, pjTimescaleWeeks)
                Else
                    Set tsvs = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineCost, pjTimescaleMonths) 'v3.4
                End If
                
                For Each tsv In tsvs

                    If tsv.Value <> "" Then

                        If tsvs.Count = 1 Then

                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                        Else

                            Select Case tsv.Index

                                Case 1

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                Case tsvs.Count

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                Case Else

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                            End Select

                        End If

                    End If

                Next tsv

                Exit Sub

            Case pjResourceTypeMaterial

                If TsvScale = "Weekly" Then 'v3.4
                    Set tsvs = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                Else
                    Set tsvs = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleMonths) 'v3.4
                End If
                
                For Each tsv In tsvs

                    If tsv.Value <> "" Then

                        If tsvs.Count = 1 Then

                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                        Else

                            Select Case tsv.Index

                                Case 1

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.BaselineStart, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                                Case tsvs.Count

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tAssign.BaselineFinish, dateFmt)

                                Case Else

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, dateFmt) & "," & Format(tsv.EndDate - 1, dateFmt)

                            End Select

                        End If

                    End If

                Next tsv

                Exit Sub

            End Select

        Case Else

            Exit Sub

    End Select

End Sub

Private Function get_Assignment_Pcnt(ByVal tAssignment As Assignment) As String

    Dim pcntField As String

    If AssignmentPCNT_Used = False Then 'v3.3.2
    
        get_Assignment_Pcnt = tAssignment.Task.GetField(FieldNameToFieldConstant(fPCNT))
        Exit Function
    
    Else
    
        pcntField = fAssignPcnt
        
    End If

    Select Case FieldNameToFieldConstant(pcntField)
    
        Case FieldNameToFieldConstant("Number1")
        
            get_Assignment_Pcnt = tAssignment.Number1
            Exit Function
        
        Case FieldNameToFieldConstant("Number2")
        
            get_Assignment_Pcnt = tAssignment.Number2
            Exit Function
        
        Case FieldNameToFieldConstant("Number3")
        
            get_Assignment_Pcnt = tAssignment.Number3
            Exit Function
        
        Case FieldNameToFieldConstant("Number4")
        
            get_Assignment_Pcnt = tAssignment.Number4
            Exit Function
        
        Case FieldNameToFieldConstant("Number5")
        
            get_Assignment_Pcnt = tAssignment.Number5
            Exit Function
        
        Case FieldNameToFieldConstant("Number6")
        
            get_Assignment_Pcnt = tAssignment.Number6
            Exit Function
        
        Case FieldNameToFieldConstant("Number7")
        
            get_Assignment_Pcnt = tAssignment.Number7
            Exit Function
        
        Case FieldNameToFieldConstant("Number8")
        
            get_Assignment_Pcnt = tAssignment.Number8
            Exit Function
        
        Case FieldNameToFieldConstant("Number9")
        
            get_Assignment_Pcnt = tAssignment.Number9
            Exit Function
        
        Case FieldNameToFieldConstant("Number10")
        
            get_Assignment_Pcnt = tAssignment.Number10
            Exit Function
            
        Case FieldNameToFieldConstant("Number11")
        
            get_Assignment_Pcnt = tAssignment.Number11
            Exit Function
        
        Case FieldNameToFieldConstant("Number12")
        
            get_Assignment_Pcnt = tAssignment.Number12
            Exit Function
        
        Case FieldNameToFieldConstant("Number13")
        
            get_Assignment_Pcnt = tAssignment.Number13
            Exit Function
        
        Case FieldNameToFieldConstant("Number14")
        
            get_Assignment_Pcnt = tAssignment.Number14
            Exit Function
        
        Case FieldNameToFieldConstant("Number15")
        
            get_Assignment_Pcnt = tAssignment.Number15
            Exit Function
        
        Case FieldNameToFieldConstant("Number16")
        
            get_Assignment_Pcnt = tAssignment.Number16
            Exit Function
        
        Case FieldNameToFieldConstant("Number17")
        
            get_Assignment_Pcnt = tAssignment.Number17
            Exit Function
        
        Case FieldNameToFieldConstant("Number18")
        
            get_Assignment_Pcnt = tAssignment.Number18
            Exit Function
        
        Case FieldNameToFieldConstant("Number19")
        
            get_Assignment_Pcnt = tAssignment.Number19
            Exit Function
        
        Case FieldNameToFieldConstant("Number20")
        
            get_Assignment_Pcnt = tAssignment.Number20
            Exit Function
        
        Case Else
        
            get_Assignment_Pcnt = 0
            Exit Function
            
    End Select
    
    get_Assignment_Pcnt = 0

End Function

Private Sub exportQuickSort(vArray As Variant, inLow As Long, inHi As Long) 'v3.4.7
  'public domain: https://stackoverflow.com/questions/152319/vba-array-sort-function
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then exportQuickSort vArray, inLow, tmpHi 'v3.4.7
  If (tmpLow < inHi) Then exportQuickSort vArray, tmpLow, inHi 'v3.4.7
End Sub

Private Function BrowseForFolder(Optional OpenAt As Variant) As Variant 'v3.4.2
    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Choose an output folder:", 0, OpenAt)
    
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0
    
    Set ShellApp = Nothing
        Select Case Mid(BrowseForFolder, 2, 1)
            Case Is = ":"
                If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
            Case Is = "\"
                If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
            Case Else
                GoTo Invalid
        End Select
    Exit Function
    
Invalid:
        noFolderSelected = True
        BrowseForFolder = False
End Function

Private Function AssignmentResumeDate(ByVal tAssign As Assignment) As Variant 'v3.4.5

    Dim tsvW As TimeScaleValues
    Dim tsvA As TimeScaleValues
    Dim i As Integer
    
    If tAssign.ResourceType <> pjResourceTypeCost Then
        'evaluate labor and material resources
        Set tsvW = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleDays)
        Set tsvA = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledActualWork, pjTimescaleDays)
    
    Else
        'evaluate cost resources
        Set tsvW = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledCost, pjTimescaleDays)
        Set tsvA = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledActualCost, pjTimescaleDays)
    
    End If

    For i = 1 To tsvA.Count
        If tsvA(i).Value = "" And tsvW(i).Value <> "" Then
            AssignmentResumeDate = tsvW(i).StartDate
            Exit Function
        End If
    Next i
    
    AssignmentResumeDate = "NA"
           
End Function

