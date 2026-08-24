Attribute VB_Name = "cptCriticalPath_bas"
'<cpt_version>v3.5.2</cpt_version>
Option Explicit
Private CritField As String 'Stores comma seperated values for each task showing which paths they are a part of
Private GroupField As String 'Stores a single value - used to group/sort tasks in final CP view
Private SubPathField As String 'v3.5.1 stores the subpath identification field
Private SubPaths As Boolean 'v3.5.1 indicates if the user wants subpaths identified
Private SubPathMax As Integer       'v3.5.1 running max of minted sub-path ids within the current path
Private SubPathClaimed As Collection 'v3.5.1 successors that have already handed their sub-path to a first driving pred
'Custom type used to store driving path vars
Type DrivingPaths
    PrevFloat As Double 'previously evaluated path float
    CurrentFloat As Double 'currently evaluated path float
End Type
Private PathCount As Integer 'number of paths to analyze
Private CurrentPath As Integer 'currently evaluated path
Private curSubPath As Integer 'v3.5.1 Tracks sub-path IDs
Private MaxPathsFound As Integer 'track the maximum number of paths found
Type NextDriver
    UID As String
    SubPath As Integer
End Type
Private NextDrivers() As NextDriver 'Array of next drivers to be analyzed
Private NextDriverCount As Integer 'count of next drivers to be analyzed
Private tDrivingPaths As DrivingPaths 'var to store DrivingPaths type
Private AnalyzedTasks As Collection 'Collection of task relationships analyzied (From UID - To UID); unique to each path analysis
'Custom type used to store Driving Task data
Type DrivingTask
    UID As String
    tFloat As Double
End Type
Private DrivingTasks() As DrivingTask 'var to store DrivingTask type
Private drivingTasksCount As Integer 'count of DrivingTasks
Public singlePath As Boolean 'cpt controlled var for limited results to a single path
Public export_to_PPT As Boolean 'cpt controlled var for controlling user notification of completed analysis
Private CustTextFields() As String 'v2.9.0 Array of custTextFields
Private CustNumFields() As String 'v2.9.0 Array of custNumFields
Private curProj As Project 'Stores active user project - not compatible with Master/Sub Architecture v2.9.0 - set as module var for cust field mapping
Private masterProj As Boolean 'v3.0.0 stores master project status of active project based on subproject count
Private subP As SubProject 'v3.0.0 used to iterate through subprojects collection
Private subPID As Integer 'v3.0.0 used to temporarily store subproject ID
Private tempproj As Project 'v3.0.0 used to temporarily reference subprojects
Private firstTask As Boolean 'v3.0.0 used to track seed task for each path
Private Const MODULE_NAME As String = "cptCriticalPath_bas"
Private userView As String
Private subProjIndexCache As Collection 'v3.5.0 cache for get_subProj_index lookups (keyed by subproject path)
Private ganttFormatOverride As Boolean 'v3.5.0

Sub DrivingPaths()
'Primary analysis module that controls analysis
'workflow through Primary, Secondary and Tertiary
'driving paths.

    'prevent spawning
    If Not cptGetUserForm("cptCriticalPath_frm") Is Nothing Then Exit Sub
    
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    Dim t As Task 'Stores initial user selected task
    Dim tdp As TaskDependency
    Dim tdps As TaskDependencies
    Dim i As Integer 'Used to iterate through Primary/Secondary/Tertiary driver arrays
    Dim analysisTaskUID As String 'Stores user selected task for recall and selection after setting final view
    
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    'Store users active project
    Set curProj = ActiveProject 'v2.9.0 get active project before displaying field selection form
    
    'v3.0.0 - check for subprojects
    masterProj = (curProj.Subprojects.Count > 1)
    
    'used to avoid code break during intial error checks
    On Error Resume Next
    
    'Validate users selected view type
    If curProj.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
        MsgBox "Please select a View with a Task Table."
        curProj = Nothing
        Exit Sub
    End If
    
    'Validate users selected window pane - select the task table if not active
    If curProj.Application.ActiveWindow.ActivePane.Index <> 1 Then
        curProj.Application.ActiveWindow.TopPane.Activate
    End If
    
    'Exit if multiple tasks are selected
    If curProj.Application.ActiveSelection.Tasks.Count > 1 Then
        MsgBox "Select a single activity only."
        curProj = Nothing
        Exit Sub
    End If
    
    'store task of activeselection
    Set t = curProj.Application.ActiveCell.Task
    
    'Check for null task rows
    If t Is Nothing Then
        MsgBox "Select a task"
        curProj = Nothing
        Exit Sub
    End If
    
    'Avoid analyzing completed tasks
    If t.PercentComplete = 100 Then
        MsgBox "Select an incomplete task"
        curProj = Nothing
        Exit Sub
    End If
    
    'Avoid analysis on summary rows
    If t.Summary = True Then
        MsgBox "Select a non-summary task"
        curProj = Nothing
        Exit Sub
    End If
    
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    'v2.9.0 Diplay Field Selection dialog
    Dim critPathFieldMapForm As cptCritPathFields_frm
    Set critPathFieldMapForm = New cptCritPathFields_frm
    
    With critPathFieldMapForm
    
        ReadCustomFields curProj
    
        For i = 1 To UBound(CustNumFields)
            .GroupField_Combobox.AddItem CustNumFields(i)
            .SubPath_Combobox.AddItem CustNumFields(i)
        Next i
        For i = 1 To UBound(CustTextFields)
            .GroupField_Combobox.AddItem CustTextFields(i)
            .PathField_Combobox.AddItem CustTextFields(i)
            .SubPath_Combobox.AddItem CustTextFields(i)
        Next i
        
        .pathCnt_txtBox.value = 3
        
        .Caption = "cptCritical Path " & cptGetVersion("cptCriticalPath_bas")
        
        With .UserView_Combobox
            .AddItem "<Default>"
            Dim v As View
            For Each v In curProj.Views
                .AddItem v.Name
            Next v
            .ListIndex = 0
        End With
        
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        
        If singlePath Then 'v2.4.2 - hide path count when only running single path
            .pathCnt_lbl.Visible = False
            .pathCnt_txtBox.Visible = False
        Else
            .pathCnt_lbl.Visible = True
            .pathCnt_txtBox.Visible = True
        End If
        
        .Show
        
        If .Tag = "cancel" Then
            Set critPathFieldMapForm = Nothing
            Set curProj = Nothing
            Exit Sub
        End If
        
        'v2.9.0 - get user field map
        CritField = .PathField_Combobox.Text
        GroupField = .GroupField_Combobox.Text
        SubPathField = .SubPath_Combobox.Text
        SubPaths = .SubPath_Checkbox.value
        PathCount = .pathCnt_txtBox.value
        userView = .UserView_Combobox.Text
        ganttFormatOverride = .ganttFormatCheckBox.value
    
    End With
    
    'Suspend calculations and screen updating
    curProj.Application.Calculation = pjManual
    curProj.Application.ScreenUpdating = False
    
    'v3.5.0 update to set custom fields in subproject schedule that align with the master project
    Dim origGroupField As Long
    Dim origCritField As Long
    Dim origSubPathField As Long
    origGroupField = FieldNameToFieldConstant(GroupField)
    origCritField = FieldNameToFieldConstant(CritField)
    origSubPathField = FieldNameToFieldConstant(SubPathField)
    
    If masterProj = True Then
        For Each subP In curProj.Subprojects
            FileOpenEx subP.Path, True
            Set tempproj = ActiveProject
            SetGroupCPFieldLookupTable FieldConstantToFieldName(origGroupField), _
                FieldConstantToFieldName(origCritField), FieldConstantToFieldName(origSubPathField), SubPaths, tempproj 'v3.5.0
        Next subP
        curProj.Activate
    End If
    
    'v3.0.0 run no matter what the masterProj condition is
    'still need to update fields in Master Project file
    'in case tasks exist at top level
    SetGroupCPFieldLookupTable GroupField, CritField, SubPathField, SubPaths, curProj
    
    'Erase previous Crit and Group field values
    CleanCritFlag curProj
    
    'Erase any previously created/modified view elements
    CleanViews curProj
    
    'Initialize Analyzed Tasks Collection
    Set AnalyzedTasks = New Collection
    
    'Add selected task to Analyzed Tasks collection and store UID for later reference
    '**NOTE** in master project scenario, will present as master project unique for selected task
    AnalyzedTasks.Add t.UniqueID, t.UniqueID & "-" & t.UniqueID
    analysisTaskUID = t.UniqueID

    'Set default Float values
    tDrivingPaths.PrevFloat = 0
    tDrivingPaths.CurrentFloat = 0
    
    'Set default driver counts
    NextDriverCount = 0
    
    '********************************
    '***Find Primary Driving Paths***
    '********************************
    
    CurrentPath = 1
    curSubPath = 1
    SubPathMax = 1
    Set SubPathClaimed = New Collection
    
    MaxPathsFound = CurrentPath
    
    'Store dependencies of user selected task
    Set tdps = t.TaskDependencies
    
    'Note that selected task is visible on path 1
    t.SetField FieldNameToFieldConstant(CritField), "1"
    t.SetField FieldNameToFieldConstant(GroupField), "1"
    If SubPaths Then t.SetField FieldNameToFieldConstant(SubPathField), "1"
    
    'Evlauate list of dependencies on selected analysis task
    For Each tdp In tdps
    
        'v3.0.0
        firstTask = True
    
        'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
        evaluateTaskDependencies tdp, t, curProj, AnalyzedTasks, curSubPath
        
    Next tdp 'Next user selected analysis task dependency
    
    '<---cpt:exit here for single driving path--->
    If singlePath Then GoTo ShowAndTell
    
    'Clear variables for re-use in evaluating secondary driver
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Set AnalyzedTasks = New Collection
    
    'Iterate through drivingtasks array to find next path driver
    FindNextDriver
    
    '**********************************
    '***Find Other Driving Paths***
    '**********************************
    
    For CurrentPath = 2 To PathCount
        
        If NextDriverCount > 0 Then
        
            MaxPathsFound = CurrentPath
            SubPathMax = NextDriverCount
            Set SubPathClaimed = New Collection
        
            'iterate through list of secondary drivers
            For i = 1 To NextDriverCount
                
                'store the current driving task
                Set t = curProj.Tasks.UniqueID(NextDrivers(i).UID)
                
                'v3.5.1 store the curSubPath
                curSubPath = NextDrivers(i).SubPath
                
                'add the driving task to the analyzed tasks collection
                AnalyzedTasks.Add t.UniqueID, t.UniqueID & "-" & t.UniqueID 'v3.5.1 - was missing key
                
                'If the task has not already been analyzed during previous path analysis,
                'set the Crit and Group Field values
                If t.GetField(FieldNameToFieldConstant(CritField)) = vbNullString Then
                    With t
                        .SetField FieldNameToFieldConstant(CritField), CurrentPath
                        .SetField FieldNameToFieldConstant(GroupField), CurrentPath
                        If SubPaths Then .SetField FieldNameToFieldConstant(SubPathField), curSubPath
                    End With
                Else
                
                    'If the task has already been analyzed during the previous path analysis,
                    'append path value to the Crit and Group Fields
                    If InStr(t.GetField(FieldNameToFieldConstant(CritField)), CurrentPath) = 0 Then
                        t.SetField FieldNameToFieldConstant(CritField), t.GetField(FieldNameToFieldConstant(CritField)) & "," & CurrentPath
                    End If
                    
                End If
                
                'Store secondary driving task dependencies
                Set tdps = t.TaskDependencies
                
                'Evlauate list of dependencies on secondary driving task
                For Each tdp In tdps
                
                    firstTask = True 'v3.5.1
                
                    'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
                    evaluateTaskDependencies tdp, t, curProj, AnalyzedTasks, curSubPath
                    
                Next tdp 'Next secondary driver dependency
                
            Next i 'next Secondary Path Driver
        
        End If
        
        'Clear variables for re-use in evaluating secondary driver
        Set tdps = Nothing
        Set tdp = Nothing
        Set t = Nothing
        Set AnalyzedTasks = New Collection
        
        tDrivingPaths.PrevFloat = tDrivingPaths.CurrentFloat
        
        'Iterate through drivingtasks array to find next path driver
        FindNextDriver
        curSubPath = 1
        
    Next CurrentPath
    
ShowAndTell:
    
    'Create and Apply the "ClearPlan Driving Path" Table, View, Group, and Filter
    SetupCPView GroupField, curProj, analysisTaskUID
    
    If Not (export_to_PPT) Then MsgBox "Complete" & vbCr & vbCr & MaxPathsFound & " path(s) identified.", vbOKOnly, "ClearPlan Critical Path Analyzer"
    
    GoTo CleanUp
    
ErrorHandler:

    Call cptHandleErr(MODULE_NAME, "cptCriticalPath", err, Erl, "Error identifying driving paths")

CleanUp:

    'Clear variables
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Erase NextDrivers, DrivingTasks
    Set AnalyzedTasks = Nothing
    NextDriverCount = 0
    drivingTasksCount = 0
    PathCount = 0
    
    'Enable calculations and screenupdating
    curProj.Application.Calculation = pjAutomatic
    curProj.Application.ScreenUpdating = True
    
    'release project variable
    Set curProj = Nothing
    Set subProjIndexCache = Nothing

End Sub

Private Sub evaluateTaskDependencies(ByVal tdp As TaskDependency, ByVal t As Task, ByVal curProj As Project, ByRef curAnalyzedTasks As Collection, ByVal curSubPath As Integer)
'Evaluate each task dependency, ignoring complete preds, then store as an analyzed relationship and evaluate criticality

    'v3.0.0 new variables
    Dim real_ToUID As Long
    Dim real_FromUID As Long
    Dim subIndex As Integer

    'v3.0.0 need to convert the
    If firstTask = True And masterProj = True Then
        firstTask = False
        If tdp.To.ExternalTask = True Then
            subIndex = get_subProj_index(curProj, tdp.To.Project)
            real_ToUID = get_tdp_MasterUID(tdp.To.UniqueID, subIndex)
        Else
            subIndex = get_subProj_index(curProj, curProj.Subprojects(tdp.To.Project).Path)
            real_ToUID = get_tdp_MasterUID(tdp.To.UniqueID, subIndex)
        End If
    Else
        real_ToUID = tdp.To.UniqueID
    End If
    
    'Only evaluate incomplete predecessors
    If real_ToUID = t.UniqueID And tdp.From.PercentComplete <> 100 Then
        'v3.0.0 account for master project condition
        If masterProj Then
        
            real_ToUID = ResolveMasterUID(curProj, tdp.To)
        
            real_FromUID = ResolveMasterUID(curProj, tdp.From)
            
        Else
        
            real_ToUID = tdp.To.UniqueID
            real_FromUID = tdp.From.UniqueID

        End If
        
        'Check dependency for existance in analyzed tasks collection
        If ExistsInCollection(curAnalyzedTasks, real_FromUID & "-" & real_ToUID) = False Then 'v3.0.0 updated with real UID for master projects
            'If dependency has not been analyzed, add to analyzed tasks collection
            curAnalyzedTasks.Add real_FromUID, real_FromUID & "-" & real_ToUID 'v3.0.0 updated with real uid for master projects
            'Calculate True Float value and evaluate against list of driving tasks
            CheckCritTask curProj, tdp, curSubPath
        End If
    End If
    
End Sub

Private Sub SetGroupCPFieldLookupTable(ByVal GroupField As String, ByVal CritField As String, ByVal SubPathField As String, ByVal SubPaths As Boolean, ByVal currentProject As Project)
'Set Crit and Group field names, assign lookup table to Group Field
    
    'v3.0.0 remove crit field attributes
    currentProject.Application.CustomFieldPropertiesEx FieldID:=FieldNameToFieldConstant(CritField), Attribute:=pjFieldAttributeNone, SummaryCalc:=pjCalcNone, GraphicalIndicators:=False, AutomaticallyRolldownToAssn:=False
    'v3.5.1 SubPath Support
    If SubPaths Then
        currentProject.Application.CustomFieldPropertiesEx FieldID:=FieldNameToFieldConstant(SubPathField), Attribute:=pjFieldAttributeNone, SummaryCalc:=pjCalcNone, GraphicalIndicators:=False, AutomaticallyRolldownToAssn:=False
    End If
    'currentProject.Application.CustomFieldRename FieldID:=FieldNameToFieldConstant(CritField), NewName:="CP Driving Paths"
    
    'Setup Lookup Table Properties
    currentProject.Application.CustomFieldPropertiesEx FieldID:=FieldNameToFieldConstant(GroupField), Attribute:=pjFieldAttributeNone
    currentProject.Application.CustomOutlineCodeEditEx FieldID:=FieldNameToFieldConstant(GroupField), OnlyLookUpTableCodes:=True, OnlyLeaves:=False, LookupDefault:=False, SortOrder:=0
    currentProject.Application.CustomFieldPropertiesEx FieldID:=FieldNameToFieldConstant(GroupField), Attribute:=pjFieldAttributeValueList, SummaryCalc:=pjCalcNone, GraphicalIndicators:=False, AutomaticallyRolldownToAssn:=False
    'currentProject.Application.CustomFieldRename FieldID:=FieldNameToFieldConstant(GroupField), NewName:="CP Driving Path Group ID"
    
    'Assign Lookup Table Values
    Dim pathNum As Integer
    
    currentProject.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), 0, "N/A"
    
    For pathNum = 1 To PathCount
        currentProject.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), pathNum, "Path " & pathNum
    Next pathNum

End Sub
Private Sub SetupCPView(ByVal GroupField As String, ByVal curProj As Project, ByVal tUID As String)
'Setup CP View with Table & Grouping by Path Value

    Dim t As Task 'used to store user selected anlaysis task
    
    If userView <> "<Default>" Then
        curProj.Application.ViewApply Name:=userView
    Else
    
        'Create CP Driving Path Table
        curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, Create:=True, ShowAddNewColumn:=True, OverwriteExisting:=True, FieldName:="ID", Width:=5, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, LockFirstColumn:=True, ColumnPosition:=0
        
        'Add fields to CP Driving Path Table
        curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Unique ID", Width:=10, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=1, LockFirstColumn:=True
        curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:=GroupField, Title:="Driving Path", Width:=5, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=1
        curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Name", Width:=45, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=2
        curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Duration", Width:=10, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=3
        curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Start", Width:=15, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=4
        curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Finish", Width:=15, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=5
        curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Total Slack", Width:=10, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=6

    End If

    'Create CP Driving Path Filter
    curProj.Application.FilterEdit Name:="*ClearPlan Driving Path Filter", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:=GroupField, test:="is greater than", value:="0", ShowInMenu:=False, ShowSummaryTasks:=False
    
    'On Error Resume Next
    
    'Create CP Driving Path Group
    Dim cpGroup As Group
    Set cpGroup = curProj.TaskGroups.Add(Name:="*ClearPlan Driving Path Group", FieldName:=GroupField)
    
    If SubPaths Then cpGroup.GroupCriteria.Add FieldName:=SubPathField, Ascending:=True
    
    'Create and apply CP Driving Path view if necessary
    If userView = "<Default>" Then
        curProj.Application.ViewEditSingle Name:="*ClearPlan Driving Path View", Create:=True, ShowInMenu:=True, Table:="*ClearPlan Driving Path Table", Filter:="*ClearPlan Driving Path Filter", Group:="*ClearPlan Driving Path Group"
        
        'Apply the CP Driving Path view
        curProj.Application.ViewApply Name:="*ClearPlan Driving Path View"
        curProj.Application.GanttBarEditEx Item:="1", RightText:="" '2.4.2 - remove resource names from view
    Else
        curProj.Application.FilterApply Name:="*ClearPlan Driving Path Filter"
        curProj.Application.GroupApply Name:="*ClearPlan Driving Path Group"
    End If
    
    'Sort the View by Finish, then by Duration to produce Waterfall Gantt
    curProj.Application.Sort Key1:="Finish", Ascending1:=True, Key2:="Duration", Ascending2:=False, Outline:=False
    
    'Select all tasks and zoom the Gantt to display all tasks in view
    curProj.Application.SelectAll
    curProj.Application.ZoomTimescale Selection:=True
    
    curProj.Application.SelectRow 1
    
    If SubPaths Then curProj.Application.SelectRow 2
    
    'Iterate through each task in view and color the Gantt bars based on CP Group Code
    If ganttFormatOverride Then
        For Each t In ActiveProject.Tasks
            If Not t Is Nothing Then 'Fix issue 44 for v2.8
            
                Dim pathValue As Integer
                
                pathValue = t.GetField(FieldNameToFieldConstant(GroupField))
            
                If pathValue = 0 Then
                    GoTo NextTask
                End If
                
                If t.Milestone = True Then
                    GoTo NextTask
                End If
            
                If masterProj Then
                    t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=StoplightColor(MaxPathsFound, pathValue), MiddleColor:=StoplightColor(MaxPathsFound, pathValue), EndColor:=StoplightColor(MaxPathsFound, pathValue), ProjectName:=curProj.Subprojects(t.Project).Path
                Else
                    t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=StoplightColor(MaxPathsFound, pathValue), MiddleColor:=StoplightColor(MaxPathsFound, pathValue), EndColor:=StoplightColor(MaxPathsFound, pathValue)
                End If
    
            End If
    
NextTask:
    
        Next t
    End If
    
    'select the users original analysis task
    curProj.Application.FindEx "UniqueID", "equals", tUID

End Sub

Private Function StoplightColor(ByVal maxValue As Long, ByVal currentValue As Long) As Long
    ' Returns a color for the given value from Red (1) => Orange => Yellow => Yellow-Green => Green (max)
    ' Returns Gray if currentValue = 0 or outside 1...maxValue
    Dim stops() As Variant
    Dim pos As Double
    Dim idxLow As Integer, idxHigh As Integer
    Dim colorLow As Variant, colorHigh As Variant
    Dim r As Long, g As Long, b As Long
    Dim fraction As Double
    
    ' Define the stoplight colors: Red, Orange, Yellow, Yellow-Green, Green
    stops = Array(RGB(179, 0, 2), _
                  RGB(240, 145, 3), _
                  RGB(240, 224, 3), _
                  RGB(140, 213, 11))
    
    If currentValue = 0 Or currentValue < 1 Or currentValue > maxValue Then
        StoplightColor = RGB(128, 128, 128)  ' Gray
        Exit Function
    End If
    
    Dim nStops As Long
    nStops = UBound(stops)                  ' Number of stops minus 1 (i.e., 4)
    
    If maxValue <> 1 Then
        pos = (currentValue - 1) / (maxValue - 1) * nStops   ' Map value to position between 0 and nStops
    Else
        'v3.5.2 if only one path, color = red
        StoplightColor = RGB(179, 0, 2)
        Exit Function
    End If
    
    idxLow = Int(pos)
    idxHigh = WorksheetFunction.Min(idxLow + 1, nStops)
    fraction = pos - idxLow
    
    colorLow = stops(idxLow)
    colorHigh = stops(idxHigh)
    
    ' Linear interpolate each RGB channel
    r = (colorLow And &HFF) + fraction * ((colorHigh And &HFF) - (colorLow And &HFF))
    g = ((colorLow \ &H100) And &HFF) + fraction * (((colorHigh \ &H100) And &HFF) - ((colorLow \ &H100) And &HFF))
    b = ((colorLow \ &H10000) And &HFF) + fraction * (((colorHigh \ &H10000) And &HFF) - ((colorLow \ &H10000) And &HFF))
    
    StoplightColor = RGB(r, g, b)
End Function

Private Sub CleanCritFlag(ByVal curProj As Project)
'Remove previous analysis values from the Crit and Group fields

    Dim t As Task 'store task var
    
    'iterate through every task in the project
    For Each t In curProj.Tasks
    
        If Not t Is Nothing Then 'Fix issue #44 for v2.8
        
            'Reset values
            t.SetField FieldNameToFieldConstant(CritField), vbNullString
            If SubPaths Then t.SetField FieldNameToFieldConstant(SubPathField), vbNullString 'v3.5.1
            'v3.0.0
            If t.Summary = False Then t.SetField FieldNameToFieldConstant(GroupField), "0"
            
        End If
    Next t

End Sub

Private Sub CleanViews(ByVal curProj As Project)
'Iterate through all Views, Tables, Filters, and Groups
'Delete previously created CP View Elements to avoid user modification errors

    Dim cpView As View
    Dim allViews As Views
    Dim cpTable As Table
    Dim allTables As Tables
    Dim cpFilter As Filter
    Dim allFilters As Filters
    Dim cpGroup As Group
    Dim allGroups As Groups
    
    'Set vars
    Set allViews = curProj.Views
    Set allTables = curProj.TaskTables
    Set allFilters = curProj.TaskFilters
    Set allGroups = curProj.TaskGroups
    
    'If the CPCritPathView is active, choose a different view
    curProj.Application.ViewApply Name:="Gantt Chart"

    On Error Resume Next
    curProj.Views("*ClearPlan Driving Path View").Delete
    curProj.TaskTables("*ClearPlan Driving Path Table").Delete
    curProj.TaskFilters("*ClearPlan Driving Path Filter").Delete
    curProj.TaskGroups("*ClearPlan Driving Path Group").Delete
    On Error GoTo 0

End Sub

Private Sub FindNextDriver()
'Iterate through Driving Tasks array to find driving tasks based on True Float value

    Dim i As Integer 'Counter used to iterate through DrivingTasks array
    Dim driverCount As Integer 'count of driving tasks found
    Dim driverFloat As Double 'float value of driving tasks

    'If no drivers were found, exit the subroutine
    If drivingTasksCount = 0 Then
        Exit Sub
    End If
    
    'Store default float and count values
    driverFloat = 0
    driverCount = 0

    'Iterate through Driving Tasks array and find the least float value
    For i = 1 To UBound(DrivingTasks)
    
        'store first float value, otherwise evaluate current float value against previously stored value
        If DrivingTasks(i).tFloat > tDrivingPaths.PrevFloat And driverFloat = 0 Then
            driverFloat = DrivingTasks(i).tFloat
        Else
            If DrivingTasks(i).tFloat > tDrivingPaths.PrevFloat And DrivingTasks(i).tFloat < driverFloat Then
                driverFloat = DrivingTasks(i).tFloat
            End If
        End If
    Next i 'Next driving task
    
    'Find all drivers with similar float and store as parallel driving tasks
    If driverFloat <> 0 Then
        For i = 1 To UBound(DrivingTasks)
            With DrivingTasks(i)
                If .tFloat = driverFloat Then
                    driverCount = driverCount + 1
                    ReDim Preserve NextDrivers(1 To driverCount)
                    NextDrivers(driverCount).UID = .UID
                    NextDrivers(driverCount).SubPath = driverCount
                End If
            End With
        Next i 'Next Driving Task
    End If
    
    'Set Tertiary Float value equal to the evaluated driving task float
    tDrivingPaths.CurrentFloat = driverFloat
    
    'set tertiary driver count
    NextDriverCount = driverCount

End Sub

Private Function FindInArray(UID As String) As Variant
'Search DrivingTasks array for a task UID

    Dim i As Long 'counter to iterate through Driving Tasks
    
    For i = LBound(DrivingTasks) To UBound(DrivingTasks)
        If DrivingTasks(i).UID = UID Then
            FindInArray = i
            Exit Function
        End If
    Next i

    FindInArray = Null

End Function

Private Sub CheckCritTask(ByVal curProj As Project, ByVal tdp As TaskDependency, ByVal curSubPath As Integer)
'Compare current task dependency against full list of Driving Tasks and
'add-to/create/replace list of Path Drivers if critical

    Dim tdps As TaskDependencies 'store task dependencies
    Dim tdpI As TaskDependency 'store task dependency
    Dim tempFloat As Double 'tempFloat value used to compare float amongst all preds
    Dim i As Variant 'used to store unique ID of driving task if found in Driving Tasks array
    Dim predT As Task 'var to store pred task of evaluated dependency relationship
    Dim succT As Task 'var to store succ task of evaluated dependency relationship
    Dim predCritCoding As String 'var to store/modify existing Crit field values
    Dim subpIndex As Integer 'v3.0.0
    Dim realPredUID As Long 'v3.0.0
    Dim realSuccUID As Long 'v3.0.0
    
    'Assign the dependency predecessor task to predT var
    'v3.0.0 consider mast project condition
    If masterProj Then
        If tdp.From.ExternalTask = True Then
            subpIndex = get_subProj_index(curProj, tdp.From.Project)
            If subpIndex = 0 Then 'subproject is not present
                Exit Sub
            Else
                realPredUID = get_external_MasterUID(tdp.From, subpIndex)
                Set predT = curProj.Tasks.UniqueID(realPredUID)
            End If
                
        Else
            subpIndex = get_subProj_index(curProj, curProj.Subprojects(tdp.From.Project).Path)
            realPredUID = get_tdp_MasterUID(tdp.From.UniqueID, subpIndex)
            Set predT = curProj.Tasks.UniqueID(realPredUID)
        End If
    Else
        realPredUID = tdp.From.UniqueID
        Set predT = curProj.Tasks.UniqueID(tdp.From.UniqueID)
    End If
    
    
    'store predecessor task Crit path coding
    predCritCoding = predT.GetField(FieldNameToFieldConstant(CritField))
    
    'Assign the dependency successor task to the succT var
    'v3.0.0 consider master project condition - succ T will never be an external task
    If masterProj Then
        subpIndex = get_subProj_index(curProj, curProj.Subprojects(tdp.To.Project).Path)
        realSuccUID = get_tdp_MasterUID(tdp.To.UniqueID, subpIndex)
        Set succT = curProj.Tasks.UniqueID(realSuccUID)
    Else
        realSuccUID = tdp.To.UniqueID
        Set succT = curProj.Tasks.UniqueID(tdp.To.UniqueID)
    End If
    
    'get the TrueFloat of Dependency relationship
    tempFloat = TrueFloat(predT, succT, tdp.Type, tdp.Lag, tdp.LagType)

    'If not evaluating the last path, and the TrueFloat value is not 0
    'Evaluate total network float and store in Driving Tasks array
    If CurrentPath < PathCount And tempFloat <> 0 Then
        
        'If other Driving Tasks have been found, Evaluate further
        i = Null
        
        If drivingTasksCount > 0 Then
        
            i = FindInArray(CStr(realPredUID))
            
            If Not IsNull(i) Then
            
                ' update existing — keep the lower float
                Dim newFloat As Double
                
                newFloat = tempFloat + tDrivingPaths.CurrentFloat
                
                If CurrentPath = 1 Then
                    If tempFloat < DrivingTasks(i).tFloat Then DrivingTasks(i).tFloat = tempFloat
                Else
                    If newFloat < DrivingTasks(i).tFloat Then DrivingTasks(i).tFloat = newFloat
                End If
                
            Else
            
                drivingTasksCount = drivingTasksCount + 1
                ReDim Preserve DrivingTasks(1 To drivingTasksCount)
                DrivingTasks(drivingTasksCount).UID = realPredUID
                DrivingTasks(drivingTasksCount).tFloat = tempFloat + tDrivingPaths.CurrentFloat
                
            End If

        Else 'No other driving tasks found, this is the first driving task
            
            'Add the new driver to the driving tasks count and store in array
            drivingTasksCount = drivingTasksCount + 1
            ReDim DrivingTasks(1 To drivingTasksCount) 'removed Preserve - should not be neccessary when finding first driving task
            DrivingTasks(drivingTasksCount).UID = realPredUID 'v3.0.0
            
            DrivingTasks(drivingTasksCount).tFloat = tempFloat + tDrivingPaths.CurrentFloat

        End If
    End If
    
    'Evaluate new driver if True Float is 0
    If tempFloat = 0 Then
        
        'If other drivers exist, and not evaluating the last path, evaluate further
        If drivingTasksCount > 0 And CurrentPath < PathCount Then 'v3.1.1
        
            'Look for predecessor task in Driving Tasks Array
            i = FindInArray(CStr(realPredUID)) 'v3.0.0
    
            'If the task exists in the driving tasks array, update the float value
            If Not IsNull(i) Then
                DrivingTasks(i).tFloat = tempFloat
            Else 'If this is a new driver
            
                'Store the driving task in the Driving Tasks array
                drivingTasksCount = drivingTasksCount + 1
                ReDim Preserve DrivingTasks(1 To drivingTasksCount)
                With DrivingTasks(drivingTasksCount)
                    .UID = realPredUID 'v3.0.0
                    .tFloat = tempFloat
                End With
            End If
            
        Else 'If no other driving tasks exists and not evaluating the last path
            If CurrentPath < PathCount Then

                'Store the new driving task
                drivingTasksCount = drivingTasksCount + 1
                ReDim DrivingTasks(1 To drivingTasksCount) 'removed Preserve - should not be neccessary when finding first driving task
                With DrivingTasks(drivingTasksCount)
                    .UID = realPredUID 'v3.0.0
                    .tFloat = tempFloat
                End With
            End If
        End If
        
        'v3.5.1 first-pred-inherits: the first driving predecessor of this successor
        'inherits the successor's sub-path (the spine keeps a stable id); each additional
        'converging predecessor mints a fresh id.
        Dim predSubPath As Integer
        If SubPaths Then
            If ExistsInCollection(SubPathClaimed, CStr(realSuccUID)) Then
                'Only mint a new sub-path id if this pred will actually receive it.
                'If it's already coded, reuse its existing id so we don't burn a number.
                If predCritCoding = vbNullString Then
                    SubPathMax = SubPathMax + 1
                    predSubPath = SubPathMax
                Else
                    predSubPath = predT.GetField(FieldNameToFieldConstant(SubPathField))
                    If predSubPath = 0 Then predSubPath = curSubPath
                End If
            Else
                predSubPath = curSubPath
                SubPathClaimed.Add realSuccUID, CStr(realSuccUID)
            End If
        End If
    
        'If no existing path coding, then no need to concatenate
        If predCritCoding = vbNullString Then
            With predT
                .SetField FieldNameToFieldConstant(CritField), CurrentPath
                .SetField FieldNameToFieldConstant(GroupField), CurrentPath
                If SubPaths Then .SetField FieldNameToFieldConstant(SubPathField), predSubPath
            End With
        Else 'if existing code, then concatenate string
            If InStr(predCritCoding, CurrentPath) = 0 Then   ' was PathCount
                predT.SetField FieldNameToFieldConstant(CritField), predCritCoding & "," & CurrentPath
            End If
        End If
    
        'store dependecies of the currently evaluted dependency
        Set tdps = predT.TaskDependencies
        
        'Iterate through the dependencies of the dependency
        For Each tdpI In tdps
        
            'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
            evaluateTaskDependencies tdpI, predT, curProj, AnalyzedTasks, predSubPath

        Next tdpI 'Next dependency of the currently evaluated dependency
    End If
        
End Sub

Private Function TrueFloat(ByVal tPred As Task, ByVal tSucc As Task, ByVal dType As Integer, ByVal dLag As Double, dlagtype As Integer) As Double
'Find True Float Value
'True Float is the dependency level 'free float' value,
'taking into consideration all duration types (including eDays),
'task calendars, leads/lags, etc

    Dim pDate As Date 'Store predecessor date (start or fin depending on link type)
    Dim sDate As Date 'Store successor date (start or fin depending on link type)
    Dim sCalObj As Calendar 'Store successor task calendar or project calendar if task cal = N/A
    Dim pCalObj As Calendar 'Store predecessor task calendar or project calendar if task cal = N/A
    Dim tempFloat As Double 'store True Float for function return
    Dim subpIndex As Integer 'v3.0.0
    
    'If pred task has a task calendar, store
    If tPred.Calendar <> "None" Then
        Set pCalObj = tPred.CalendarObject
    Else 'If no task calendar, store project cal
        'v3.0.0 consider master project condition
        If masterProj = True Then
            If tPred.Project = curProj.Tasks.UniqueID(0).Project Then 'task is in master project
                Set pCalObj = curProj.Calendar
            Else
                Set pCalObj = curProj.Subprojects(tPred.Project).SourceProject.Calendar
            End If
        Else
            Set pCalObj = ActiveProject.Calendar
        End If
    End If
    
    'If succ task has a task calendar, store
    If tSucc.Calendar <> "None" Then
        Set sCalObj = tSucc.CalendarObject
    Else 'If no task calendar, store project cal
        'v3.0.0 consider master project condition
        If masterProj = True Then
            If tSucc.Project = curProj.Tasks.UniqueID(0).Project Then 'task is in master project
                Set sCalObj = curProj.Calendar
            Else
                Set sCalObj = curProj.Subprojects(tSucc.Project).SourceProject.Calendar
            End If
        Else
            Set sCalObj = ActiveProject.Calendar
        End If
    End If
    
    ' Determine pred and succ reference dates based on dependency type
    Dim pDateBase As Date, sDateBase As Date, sDateEarly As Date
    
    Select Case dType
    
        Case 0: pDateBase = tPred.Finish: sDateBase = tSucc.Finish: sDateEarly = tSucc.EarlyFinish  'FF
        
        Case 1: pDateBase = tPred.Finish: sDateBase = tSucc.Start:  sDateEarly = tSucc.EarlyStart   'FS
        
        Case 2: pDateBase = tPred.Start:  sDateBase = tSucc.Finish: sDateEarly = tSucc.EarlyFinish  'SF
        
        Case 3: pDateBase = tPred.Start:  sDateBase = tSucc.Start:  sDateEarly = tSucc.EarlyStart   'SS
        
        Case Else: TrueFloat = 0: Exit Function
        
    End Select
    
    ' Resolve effective succ date (leveling delay ? use early date)
    Dim sDateEffective As Date
    
    If tSucc.LevelingDelay > 0 Then
        sDateEffective = sDateEarly
    Else
        sDateEffective = sDateBase
    End If
    
    ' Apply lag or lead
    If dLag >= 0 Then
        pDate = Application.DateAdd(pDateBase, Application.DurationFormat(dLag, dlagtype), sCalObj)
        sDate = sDateEffective
    Else
        pDate = pDateBase
        sDate = Application.DateAdd(sDateEffective, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
    End If
    
    'v2.8.2 check for edays
    If Left(GetLettersOnly(tPred.DurationText), 1) <> "e" Then
    
        'no edays; subtract the pred date from the succ date, using the pred calendar, to get the True Float value
        tempFloat = Application.DateDifference(pDate, sDate, pCalObj)
        
    Else
    
        'using edays; calculate date diff in minutes
        tempFloat = DateDiff("n", pDate, sDate)
    
    End If
    
    'Return the True Float value
    TrueFloat = tempFloat

End Function

Public Function ExistsInCollection(ByVal col As Collection, ByVal vKey As Variant) As Boolean
'Check for task dependency relationship in the analyzed tasks collection

    'If error encountered, value does not exist in the collection
    On Error GoTo err
    
    col.Item vKey 'Store found item; if not found, will produce error
    ExistsInCollection = True 'Set True
    Exit Function
    
err: 'If error encountered, item does not exist - return "False" boolean vlaue
    ExistsInCollection = False
    
End Function

Function GetLettersOnly(str As String) As String
'v2.8.2 - strip out non-alpha characters from input string
'used to evaluate task duration text for elapsed day prefix "e"

    Dim i As Long, letters As String, letter As String

    letters = vbNullString

    For i = 1 To Len(str)
        letter = VBA.Mid$(str, i, 1)

        If Asc(LCase(letter)) >= 97 And Asc(LCase(letter)) <= 122 Then
            letters = letters + letter
        End If
    Next
    GetLettersOnly = letters
End Function

Private Sub ReadCustomFields(ByVal curProj As Project)
'v2.9.0 - added to allow user selection of custom fields

    Dim i As Integer

    'Read local Custom Text Fields
    ReDim CustTextFields(1 To 30)
    For i = 1 To 30
    
        If Len(curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Text" & i))) > 0 Then
            CustTextFields(i) = curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Text" & i))
        Else
            CustTextFields(i) = "Text" & i
        End If
        
    Next i

    'Read local Custom Number Fields
    ReDim CustNumFields(1 To 20)
    For i = 1 To 20

        If Len(curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Number" & i))) > 0 Then
            CustNumFields(i) = curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Number" & i))
        Else
            CustNumFields(i) = "Number" & i
        End If

    Next i

End Sub

Private Function ResolveMasterUID(ByVal proj As Project, ByVal depTask As Task) As Long
    Dim idx As Integer
    If depTask.ExternalTask Then
        idx = get_subProj_index(proj, depTask.Project)
        ResolveMasterUID = get_tdp_MasterUID(depTask.UniqueID, idx)
    Else
        idx = get_subProj_index(proj, proj.Subprojects(depTask.Project).Path)
        ResolveMasterUID = get_tdp_MasterUID(depTask.UniqueID, idx)
    End If
End Function


Function get_subProj_index(ByVal mProj As Project, ByVal subprojectFilename As String) As Integer
'v3.5.0 Returns the subproject offset ID used to calculate the displayed Master Project UID.
'In a master project, a task's displayed UID = localUID + 4194304 * (subP_Index + 1).
'We derive subP_Index by reading an actual master-context task UID for the target
'subproject and integer-dividing by 4194304, then subtracting 1 so the returned value
'plugs directly into the existing get_tdp_MasterUID / get_external_MasterUID formulas.

    Dim cached As Variant
    Dim subP As SubProject
    Dim found As Boolean
    Dim t As Task
    Dim resolvedPath As String

    'Return cached value if previously resolved for this subproject path
    If subProjIndexCache Is Nothing Then
        Set subProjIndexCache = New Collection
    Else
        On Error Resume Next
        cached = subProjIndexCache.Item(subprojectFilename)
        If err.Number = 0 Then
            On Error GoTo 0
            get_subProj_index = CInt(cached)
            Exit Function
        End If
        err.Clear
        On Error GoTo 0
    End If

    'Confirm the requested subproject exists in master
    found = False
    For Each subP In mProj.Subprojects
        If subP.Path = subprojectFilename Then
            found = True
            Exit For
        End If
    Next subP
    If Not found Then
        get_subProj_index = 0
        Exit Function
    End If

    'Find a master-context task belonging to this subproject; derive offset from its UID
    For Each t In mProj.Tasks
        If Not t Is Nothing Then
            If t.UniqueID >= 4194304 Then
                resolvedPath = vbNullString
                On Error Resume Next
                resolvedPath = mProj.Subprojects(t.Project).Path
                On Error GoTo 0
                If resolvedPath = subprojectFilename Then
                    get_subProj_index = CInt(t.UniqueID \ 4194304)
                    subProjIndexCache.Add get_subProj_index, subprojectFilename
                    Exit Function
                End If
            End If
        End If
    Next t

    'No suitable task found (e.g., empty subproject) — fall back to 0
    get_subProj_index = 0

End Function

Function get_tdp_MasterUID(ByVal subP_UID As Long, ByVal subP_Index As Integer) As Long
'v3.0.0 convert subproject format UID to master project uid format
    
    If subP_Index = 0 Then
        get_tdp_MasterUID = subP_UID
    Else
        get_tdp_MasterUID = subP_UID + 4194304 * subP_Index
    End If
    Exit Function
    
End Function

Function get_external_MasterUID(ByVal subP_Task As Task, ByVal subP_Index As Integer) As Long
'v3.0.0 get corresponding subproject UID for external reference task

    get_external_MasterUID = subP_Task.GetField(185073906) Mod 4194304 + 4194304 * subP_Index
    Exit Function

End Function
