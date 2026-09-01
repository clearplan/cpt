Attribute VB_Name = "SubprojectConsolidationTool"
Public external_link_array() As String 'Pred UID, Link Type/Lag, Succ UID
Public new_ID() As String 'UID, Row ID
Public original_task_data() As String 'UID, Start, Finish
Public ProgressWindow As SCT_ProgressForm
Public TotalTaskCount As Integer
Public dqErrorMessage As String

Sub SCT()

'Version 1.1
'Required ProjectGlobal references: Microsoft Excel xx.x Object Library, Microsoft Scripting Runtime

'**Change Log**
'v1.0.0 - accounts for SS, SF, FF, FS and + / - characters in subproject filenames when capturing cross project links
'v0.3.5 - Added "DoEvents" command to avoid system errors while attempting to open/close/save subprojects and consolidated output
'v0.3.6 - Added "DoEvents" prior to openings subprojects, changed "End" command to "Exit Sub" command in SCT subroutine
'v0.3.7 - Added error handling to avoid issues while saving
'v0.4 - Added handle for X button close on SCT form
'v1.0 - Added field pick-list
'v1.1 - Added data quality checks: summary links, external subtasks in the master, missing Project UIDs
'     - Updated calendar checks to account for null task rows i.e. blank inserts where the task object = Nothing

'This tool consolidates all subprojects within the active master project shell into a single integrated project file.
'All task dates, metadata, resources, task assignments and baseline information are translated and reviewed against the
'original documents.  An error log is produced that notes and descrepencies between the original subprojects and the
'final consolidated project.

    'check for an active master project
    Dim activemaster As Project

    If Application.Projects.Count < 1 Then
        MsgBox "Please open a Master Project with more than 1 inserted subproject."
        Exit Sub
    End If
    
    Set activemaster = ActiveProject
    
    If activemaster.Subprojects.Count < 2 Then
    
        MsgBox "Please open a Master Project with more than 1 inserted subproject."
        
        Exit Sub
        
    End If

    Dim SCTgui As SCT_Form
    Dim sct_v As String
    Dim saveLocation As String
    Dim resourceFlag As Boolean
    Dim FUID As String
    Dim ResErrors As Integer
    
    sct_v = "SCT v1.0"
    
    Set SCTgui = New SCT_Form
    
    SCTgui.Caption = sct_v
    
    SCTgui.Show
    
    Select Case SCTgui.Tag
    
        Case False
            
            Set SCTgui = Nothing
            Exit Sub
            
        Case Else
        
            saveLocation = SCTgui.FileSaveTextBox.Text
            resourceFlag = SCTgui.ResourceCheckbox.Value
            FUID = SCTgui.FileUIDCombobox.Value
            Unload SCTgui
            
    End Select
    
    'Turn off automatic calculations, suspend application alerts, expand all outline levels
    Calculation = pjManual
    Alerts (False)
    ScreenUpdating = False
    EnableEvents = False
    GroupApply ("No Group")
    SelectAll
    OutlineShowAllTasks
    
    'Count all Subtasks in All Projects
    SelectAll
    TotalTaskCount = ActiveSelection.Tasks.Count
    
    'Display progress bar and set defaults
    Set ProgressWindow = New SCT_ProgressForm
    
    With ProgressWindow
    
        .Caption = "SCT Progress"
        .ProgLabel = "Preparing Analysis... 0%"
        .ProgBar.Min = 0
        .ProgBar.Max = 100
        .ProgBar.Value = 0
    
    End With
    
    ProgressWindow.Show
    
    'Call the ReviewSubprojectCalendars function and send the active master
    Dim calError As Boolean
    calError = False
    calError = ReviewSubprojectCalendars(activemaster)
    If calError = True Then
        GoTo CalendarError
    End If
    
    'additional data quality evaluations
    Dim dqError As Boolean
    dqError = False
    dqError = DataQualityChecks(activemaster, FUID)
    If dqError = True Then
        GoTo DataQualityError
    End If
    
    On Error GoTo ErrorHandler
    
    '******DEBUG******
    'On Error GoTo 0
    '******DEBUG******
    
    'Call the PopulateArray function and send the active master project
    Call PopulateExternalLinkArray(activemaster, FUID)
    
    'Call the UnlinkSubProjects function and send the active master project
    Call UnlinkSubProjects(activemaster)
    
    'update task count to account for un-linked external tasks
    SelectAll
    TotalTaskCount = ActiveSelection.Tasks.Count
    
    'Call the CleanUpExtnernalLinks function and send the active master project
    Call CleanUpExternalLinks(activemaster, FUID)
    
    'Call the ReacreateExternalLinks function and send the active master project
    Call RecreateExternalLinks(activemaster)
    
    'If the "Remove Resources" check-box = True, call the RemoveResources function
    'and send the active master project
    If resourceFlag = True Then
        Call RemoveResources(masterproj, FUID)
    End If
    
    'Run Error Checks **NEEDS TO BE REVIEWED FOR SPEED ISSUES**
    'ResErrors = ProduceErrorReport(masterproj, FUID)
    
    Call update_progressbar("Saving Master File...", 100)
    
    'Save the consolidated master to the initially selected location
    DoEvents
    activemaster.SaveAs Name:=saveLocation

ExitMacro:
    
    DoEvents
     
    'Turn on automatic calculations and application alerts
    Calculation = pjAutomatic
    Alerts (True)
    ScreenUpdating = True
    EnableEvents = True
    Set activemaster = Nothing
    
    Unload ProgressWindow
    
    'If ResErrors = 0 Then
        MsgBox "Complete"
        Exit Sub
    'Else
        'MsgBox ResErrors & " errors were found." & vbCrLf & vbCrLf _
            '& "Please review the error log saved to your desktop."
    'End If
    
CalendarError:

    If calError = True Then
        Calculation = pjAutomatic
        Alerts (True)
        ScreenUpdating = True
        EnableEvents = True
        Set activemaster = Nothing
        
        Unload ProgressWindow
        
        MsgBox "One or more subprojects has a unique Project Calendar. " & vbCr & vbCr & _
            "Please update the task calendars and run the SCT again."
        
        Exit Sub
        
    End If
    
DataQualityError:

    If dqError = True Then
    
        Calculation = pjAutomatic
        Alerts (True)
        ScreenUpdating = True
        EnableEvents = True
        Set activemaster = Nothing
        
        Unload ProgressWindow
        
        MsgBox "There are one or more data quality errors to review: " & _
            dqErrorMessage & vbCr & vbCr & _
            "Please correct these errors and run the SCT again."
            
        Exit Sub
    
    End If
    
ErrorHandler:

    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
    
    Resume ExitMacro
    
End Sub

Function DataQualityChecks(ByVal masterproj As Project, FUID As String) As Boolean
'This function reviews the collection of subprojects for any potential data quality errors
'before attempting to complete the consolidation and returns a "True" value if any
'errors are found.

    Dim sproj As Subproject
    Dim cursproj As Project
    Dim ts As Task
    Dim cntr As Integer
    Dim externalTasksFound, summaryLinksFound, missingPUIDsFound As Boolean
    Dim sprojCntr As Integer
    
    cntr = 0
    
    dqErrorMessage = ""
    
    For Each sproj In masterproj.Subprojects
        
        Set cursproj = sproj.SourceProject
        
        For Each ts In cursproj.Tasks
        
            If Not ts Is Nothing Then
        
                cntr = cntr + 1
                
                Call update_progressbar("Performing Data Quality Checks...", CInt((cntr / TotalTaskCount) * 100))
            
                'check for summary links
                If summaryLinksFound = False And ts.Summary = True And (ts.Predecessors <> "" Or ts.Successors <> "") Then
                    summaryLinksFound = True
                    dqErrorMessage = dqErrorMessage & vbCr & vbCr & "     * There are Summary Tasks with Predecessors and/or Successors"
                    If externalTasksFound And missingPUIDsFound Then GoTo ExitDQFunction
                    GoTo NextTask
                End If
                
                'check for external tasks
                If externalTasksFound = False And ts.ExternalTask = True Then
                    For sprojCntr = 1 To masterproj.Subprojects.Count
                        If ts.Project = masterproj.Subprojects(sprojCntr).Path Then GoTo ExternalCheckExit
                    Next sprojCntr
                    externalTasksFound = True
                    dqErrorMessage = dqErrorMessage & vbCr & vbCr & "     * There are External Tasks in the Master Project indicating a linked Subproject is missing"
                    If summaryLinksFound And missingPUIDsFound Then GoTo ExitDQFunction
                    GoTo NextTask
ExternalCheckExit:
                End If
                
                'check for missing Project UIDs
                If missingPUIDsFound = False And ts.GetField(FieldNameToFieldConstant(FUID)) = "" And (InStr(ts.Predecessors, "\") > 0 Or InStr(ts.Successors, "\") > 0) Then
                    missingPUIDsFound = True
                    dqErrorMessage = dqErrorMessage & vbCr & vbCr & "     * There are missing Project UIDs on tasks with external links"
                    If externalTasksFound And summaryLinksFound Then GoTo ExitDQFunction
                    GoTo NextTask
                End If
                
            End If
NextTask:
        Next ts
        
    Next sproj
 
ExitDQFunction:
 
    If summaryLinksFound Or externalTasksFound Or missingPUIDsFound Then DataQualityChecks = True

End Function

Function ReviewSubprojectCalendars(ByVal masterproj As Project) As Boolean
'This funciton reviews the project calendars of each subproject and returns a "True" value
'if any of the subprojects has a Unique Calendar and the tasks have not been given task calendars
    
    
    Dim sproj As Subproject
    Dim cursproj As Project
    Dim ts As Task
    Dim cntr As Integer
    
    cntr = 0
    
    For Each sproj In masterproj.Subprojects
    
        Set cursproj = sproj.SourceProject
        
        For Each ts In cursproj.Tasks
        
            If Not ts Is Nothing Then
        
                        If ts.Summary = False And ts.ExternalTask = False Then
                
                    cntr = cntr + 1
                    
                    Call update_progressbar("Reviewing Subproject Calendars: " & vbCrLf & "Task " & cntr & " of " & TotalTaskCount & "... ", CInt((cntr / TotalTaskCount) * 100))
                    
                    If sproj.SourceProject.Calendar <> masterproj.Calendar Then
                    
                        If ts.Calendar <> sproj.SourceProject.Calendar Then
                            
                            ReviewSubprojectCalendars = True
                            
                            GoTo CalErrorExitFunction
                        
                        End If
                    
                    End If
                    
                End If
                
            End If

        Next ts
        
    Next sproj
    
CalErrorExitFunction:

End Function
Private Sub update_progressbar(ByVal StatusMsg As String, ByVal progress As Integer)
'This subroutine receives a status message string and a progress bar value from the calling module
'The progress window is then updated accordingly.

    With ProgressWindow
    
        If progress > .ProgBar.Max Then
            .ProgLabel = StatusMsg & .ProgBar.Max
            .ProgBar.Value = .ProgBar.Max
        Else
            .ProgLabel = StatusMsg & progress & "%"
            .ProgBar.Value = progress
        End If
        
        DoEvents
    
    End With

End Sub
Function ProduceErrorReport(ByVal masterproj As Project, ByVal FileUID As String) As Integer
'This fucntion produces an error report by comparing the original start and finish dates for each task
'against the resulting consolidated master file.  Any deltas are reported in a csv file which is saved
'to the users desktop
'***********************************************************************************************************
'**this function is incredibly slow and has been removed from the main module pending further optimizaiton**
'***********************************************************************************************************
    
    Dim fso As New FileSystemObject
    Dim outputstream As TextStream
    Dim errorcount As Integer
    Dim taskError As Boolean
    Dim delta As Boolean
    Dim t As Task
    Dim i As Integer
    
    delta = False
    taskError = False

    For Each t In masterproj.Tasks
    
        Call update_progressbar("Running Error Checks: " & vbCrLf & "Task " & t.ID & " of " & TotalTaskCount & "... ", CInt((t.ID / TotalTaskCount) * 100))
    
        If Not t Is Nothing Then
        
            If t.Summary = False Then
            
                For i = LBound(original_task_data, 2) To UBound(original_task_data, 2)
            
                    If original_task_data(0, i) = t.GetField(FieldNameToFieldConstant(FileUID)) Then
                        
                        taskError = False
                    
                        If t.Start <> original_task_data(1, i) Then
                            If delta = False Then
                            
                                Set outputstream = fso.CreateTextFile(CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator & "Consolidation Resource Errors.csv", True)
                                
                                outputstream.WriteLine "FIle UID, Original Start, New Start, Original Finish, New Finish"

                                delta = True
                            
                            End If
                            
                            If delta = True Then
                                
                                outputstream.Write t.GetField(FieldNameToFieldConstant(FileUID)) & ", " & original_task_data(1, i) & ", " & t.Start
                                
                                taskError = True
                            
                            End If
                        End If
                        
                        If t.Finish <> original_task_data(2, i) Then
                        
                            If delta = False Then
                            
                                Set outputstream = fso.CreateTextFile(CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator & "Consolidation Resource Errors.csv", True)
                                
                                outputstream.WriteLine "FIle UID, Original Start, New Start, Original Finish, New Finish"
                                
                                delta = True
                                
                            End If
                            If taskError = False Then
                                
                                outputstream.Write t.GetField(FieldNameToFieldConstant(FileUID)) & ",,, " & t.Finish & " <> " & tFinish
                                
                                taskError = True
                            
                            Else
                                
                                outputstream.Write ", " & t.Finish & ", " & original_task_data(2, i)
                            
                            End If
                        End If

                        If taskError = True Then
                            
                            outputstream.Write vbCrLf
                            
                            errorcount = errorcount + 1
                            
                        End If
                        
                        GoTo Next_Task
                                
                    End If
            
                Next i
                
            End If
        
        End If
        
Next_Task:

    Next t
    
    If delta = True Then
        outputstream.Close
        Set outputstream = Nothing
        Set fso = Nothing
    End If
    
    ProduceErrorReport = errorcount
    
End Function
Function RemoveResources(ByVal masterproj As Project, ByVal FileUID As String)
'This function removes all resource usage and task assignments from the consolidated master project.
'This function is dependent upon the value of the "Remove Resources" checkbox on the SCT Gui.
    
    Dim t As Task
    Dim a As Assignment
    
    For Each t In masterproj.Tasks

        Call update_progressbar("Removing Resources:" & vbCrLf & "Task " & t.ID & " of " & TotalTaskCount & "... ", CInt((t.ID / TotalTaskCount) * 100))
        
        If t.Summary = False Then
        
            If Not t Is Nothing Then
        
                If t.ResourceNames <> "" Then
                    
                    t.Type = pjFixedDuration
                    t.EffortDriven = False
                    
                    For Each a In t.Assignments
                        a.Delete
                    Next a
                
                    t.Work = 0
                    
                End If
            End If
        End If
    Next t

End Function

Function RecreateExternalLinks(ByVal masterproj As Project)
'This function compares the external_link_array to the new_ID array and recreates the external links as
'internal links in the newly consolidated master project.

    Dim i1 As Integer
    Dim predID As String
    Dim succID As String
    Dim t As Task
    Dim i2 As Integer
    Dim foundPred As Boolean
    Dim foundSucc As Boolean
    
    For i1 = LBound(external_link_array, 2) To UBound(external_link_array, 2)
       
        Call update_progressbar("Recreating External Links:" & vbCrLf & "Task " & (i1 + 1) & " of " & (UBound(external_link_array, 2) + 1) & "... ", CInt(((i1 + 1) / (UBound(external_link_array, 2) + 1)) * 100))
       
        foundPred = False
        foundSucc = False
        
        For i2 = LBound(new_ID, 2) To UBound(new_ID, 2)
            If new_ID(0, i2) = external_link_array(0, i1) Then
                predID = new_ID(1, i2)
                foundPred = True
            End If
            
            If new_ID(0, i2) = external_link_array(2, i1) Then
                succID = new_ID(1, i2)
                foundSucc = True
            End If
            
            If foundPred = True And foundSucc = True Then
                Set t = masterproj.Tasks(CInt(succID))
                
                If t.Predecessors = "" Then
                    t.Predecessors = t.Predecessors & predID & external_link_array(1, i1)
                Else
                    t.Predecessors = t.Predecessors & "," & predID & external_link_array(1, i1)
                End If
                
                predID = ""
                succID = ""
                
                Exit For
            
            End If
                
        Next i2
        
        Set t = Nothing
        
    Next i1

End Function

Function CleanUpExternalLinks(ByVal masterproj As Project, ByVal UIDField As String)
'This function scans through the 'flat' consolidated project and removes dummy external links.  It also
'populates the new_ID array based to use when recreating the external links.
    
    Dim t As Task
    Dim counter As Integer
    
    ReDim new_ID(0 To 1, 0)
    counter = 0
    
    For Each t In masterproj.Tasks
    
        counter = counter + 1
    
        Call update_progressbar("Cleaning up Consolidated Master:" & vbCrLf & "Task " & counter & " of " & TotalTaskCount & "... ", CInt((counter / TotalTaskCount) * 100))
        
        If Not t Is Nothing Then
            If t.Summary = False Then
                If t.ExternalTask = True Then
                    t.Delete
                    GoTo Next_Task:
                End If
                
                new_ID(0, UBound(new_ID, 2)) = t.GetField(FieldNameToFieldConstant(UIDField))
                
                new_ID(1, UBound(new_ID, 2)) = t.ID
                
                ReDim Preserve new_ID(0 To 1, UBound(new_ID, 2) + 1)
            End If
        End If
        
Next_Task:

    Next t

    ReDim Preserve new_ID(0 To 1, UBound(new_ID, 2) - 1)

End Function
Private Sub UnlinkSubProjects(ByVal masterproj As Project)
'This function cycles through each subproject in the active master project.
'Each subproject is opened and the "Link to Subproject" flag is removed to make the tasks static.

    Dim sproj As Subproject
    Dim all_sProj As Subprojects
    Dim sproj_File As String
    Dim t As Task
    Dim progressString As String

    Set all_sProj = masterproj.Subprojects
    
    For Each sproj In all_sProj
        
        On Error GoTo ErrorHandler
    
        Call update_progressbar("Unlinking Subprojects:" & vbCrLf & sproj.SourceProject.Name & ", " & sproj.Index & " of " & all_sProj.Count & "... ", CInt((sproj.Index / all_sProj.Count) * 100))

        sproj_File = sproj.SourceProject.FullName
    
        DoEvents
        
        progressString = "none"
        
NotOpened:

        FileOpen (sproj_File)
        progressString = "open"
        
        DoEvents
        
NotUnlinked:

        sproj.LinkToSource = False
        progressString = "unlinked"
        
        DoEvents

NotClosed:

        FileClose pjDoNotSave
        progressString = "closed"
        
        DoEvents
        
NextSproj:

    Next sproj
    
    Set all_sProj = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox Err.Number & ": " & Err.Description, vbOKOnly, "Error Unlinking Subprojects"
    
    Select Case progressString
        
        Case "none"
        
            DoEvents
            Resume NotOpened
            
        Case "open"
            
            DoEvents
            Resume NotUnlinked
            
        Case "unlinked"
            
            DoEvents
            Resume NotClosed
            
        Case "closed"
        
            DoEvents
            Resume NextSproj
            
        Case Else
        
            DoEvents
            Resume NextSproj
    
    End Select

End Sub

Function PopulateExternalLinkArray(ByVal masterproj As Project, ByVal UIDField As String)
'This function reads the external links and populates the external link array.
'It also copies the original task data to use when running the final error report.

    Dim t As Task
    Dim str1 As String
    Dim tempPred As String
    Dim tempPath As String
    Dim tempID As String
    Dim tempLinkInfo As String
    Dim evalLinkInfo As String
    Dim p1 As Integer
    Dim counter As Integer
    
    ReDim external_link_array(0 To 2, 0)
    'ReDim original_task_data(0 To 2, 0)
    
    counter = 0
    
    
    For Each t In masterproj.Tasks
        
        counter = counter + 1
        
        Call update_progressbar("Reviewing External Links:" & vbCrLf & "Task " & counter & " of " & TotalTaskCount & "... ", CInt((counter / TotalTaskCount) * 100))
    
        If Not t Is Nothing Then
    
            If t.Summary = False And t.ExternalTask = False Then
            
                'original_task_data(0, UBound(originaltaskdata, 2)) = t.GetField(FieldNameToFieldConstant(UIDField)) **Removed pending review & optimization of Error Log **
                'original_task_data(1, UBound(originaltaskdata, 2)) = t.Start
                'original_task_data(2, UBound(originaltaskdata, 2)) = t.Finish
                
                'ReDim Preserve original_task_data(0 To 2, UBound(original_task_data, 2) + 1)
            
                If t.TaskDependencies.Count > 0 Then
                
                    If InStr(t.Predecessors, "\") > 0 Then
                            
                        str1 = t.Predecessors
                        
                        While InStr(1, str1, ",") <> 0
                            p1 = InStr(1, str1, ",")
                            tempPred = Mid(str1, 1, p1 - 1)
                            pcount = pcount + 1
                            If InStr(tempPred, "\") > 0 Then
                                tempPath = Left(tempPred, InStrRev(tempPred, "\") - 1)
                                tempID = Right(tempPred, Len(tempPred) - Len(tempPath) - 1)
                                evalLinkInfo = tempID
                                Select Case True
                                    Case InStr(evalLinkInfo, "FF") > 0
                                        tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "FF") - 1)
                                        tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                    Case InStr(evalLinkInfo, "SS") > 0
                                        tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "SS") - 1)
                                        tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                    Case InStr(evalLinkInfo, "SF") > 0
                                        tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "SF") - 1)
                                        tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                    Case InStr(evalLinkInfo, "+") > 0
                                        tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "+") - 3)
                                        tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                    Case InStr(evalLinkInfo, "-") > 0
                                        tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "-") - 3)
                                        tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                    Case Else
                                        tempID = evalLinkInfo
                                        tempLinkInfo = ""
                                End Select
                                
                                external_link_array(0, UBound(external_link_array, 2)) = Projects(tempPath).Tasks(CInt(tempID)).GetField(FieldNameToFieldConstant(UIDField))
                                external_link_array(1, UBound(external_link_array, 2)) = tempLinkInfo
                                external_link_array(2, UBound(external_link_array, 2)) = t.GetField(FieldNameToFieldConstant(UIDField))
                                ReDim Preserve external_link_array(0 To 2, UBound(external_link_array, 2) + 1)
                                
                            End If
                            str1 = Mid(str1, p1 + 1)
                        Wend
                        tempPred = str1
                        
                        If InStr(tempPred, "\") > 0 Then
                            tempPath = Left(tempPred, InStrRev(tempPred, "\") - 1)
                            tempID = Right(tempPred, Len(tempPred) - Len(tempPath) - 1)
                            evalLinkInfo = tempID
                            Select Case True
                                Case InStr(evalLinkInfo, "FF") > 0
                                    tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "FF") - 1)
                                    tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                Case InStr(evalLinkInfo, "SS") > 0
                                    tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "SS") - 1)
                                    tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                Case InStr(evalLinkInfo, "SF") > 0
                                    tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "SF") - 1)
                                    tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                Case InStr(evalLinkInfo, "+") > 0
                                    tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "+") - 3)
                                    tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                Case InStr(evalLinkInfo, "-") > 0
                                    tempID = Left(evalLinkInfo, InStr(evalLinkInfo, "-") - 3)
                                    tempLinkInfo = Right(evalLinkInfo, Len(evalLinkInfo) - Len(tempID))
                                Case Else
                                    tempID = evalLinkInfo
                                    tempLinkInfo = ""
                            End Select
                            external_link_array(0, UBound(external_link_array, 2)) = Projects(tempPath).Tasks(CInt(tempID)).GetField(FieldNameToFieldConstant(UIDField))
                            external_link_array(1, UBound(external_link_array, 2)) = tempLinkInfo
                            external_link_array(2, UBound(external_link_array, 2)) = t.GetField(FieldNameToFieldConstant(UIDField))
                            ReDim Preserve external_link_array(0 To 2, UBound(external_link_array, 2) + 1)
                            
                        End If
                        
                        tempPred = ""
                        tempID = ""
                        tempLinkInfo = ""
                        evalLinkInfo = ""
                        str1 = ""
                        p1 = 0
                        
                    End If
                
                End If
            
            End If
            
        End If
        
    Next t
    
    ReDim Preserve external_link_array(0 To 2, UBound(external_link_array, 2) - 1)
    'ReDim Preserve original_task_data(0 To 2, UBound(original_task_data, 2) - 1)
    
End Function
