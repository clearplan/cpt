Attribute VB_Name = "cptResourceDemand_bas"
'<cpt_version>v1.6.0</cpt_version>
Option Explicit
Private Const MODULE_NAME = "cptResourceDemand_bas"

Sub cptExportResourceDemand(ByRef myResourceDemand_frm As cptResourceDemand_frm, Optional lngTaskCount As Long)
6     'objects
7     Dim oEstimates As Scripting.Dictionary
8     Dim oCalendar As MSProject.Calendar
9     Dim oRecordset As ADODB.Recordset
10    Dim oException As MSProject.Exception
11    Dim oSettings As Object
12    Dim oListObject As Excel.ListObject
13    Dim oSubproject As MSProject.SubProject
14    Dim oTask As MSProject.Task
15    Dim oResource As MSProject.Resource
16    Dim oAssignment As MSProject.Assignment
17    Dim oTSV As TimeScaleValue
18    Dim oTSVS_BCWS As TimeScaleValues
19    Dim oTSVS_WORK As TimeScaleValues
20    Dim oTSVS_AW As TimeScaleValues
21    Dim oTSVS_COST As TimeScaleValues
22    Dim oTSVS_AC As TimeScaleValues
23    Dim oCostRateTable As CostRateTable
24    Dim oPayRate As PayRate
25    Dim oExcel As Excel.Application 'Object
26    Dim oWorksheet As Excel.Worksheet 'Object
27    Dim oWorkbook As Excel.Workbook 'Object
28    Dim oRange As Excel.Range 'Object
29    Dim oPivotTable As Excel.PivotTable 'Object
30    Dim oPivotChartTable As Excel.PivotTable
31    Dim oChart As Excel.Chart
32    'dates
33    Dim dtEndDate As Date
34    Dim dtStartDate As Date
35    Dim dtWeek As Date
36    Dim dtStart As Date
37    Dim dtFinish As Date
38    'doubles
39    Dim dblWork As Double
40    Dim dblCost As Double
41    'strings
42    Dim strCFN As String
43    Dim strTask As String
44    Dim strFields As String
45    Dim strRateSets As String
46    Dim strMsg As String
47    Dim strSettings As String
48    Dim strKey As String
49    Dim strView As String
50    Dim strFileName As String
51    Dim strRange As String
52    Dim strHeader As String
53    'longs
54    Dim lngLastCol As Long
55    Dim lngItem As Long
56    Dim lngCols As Long
57    Dim lngLastRow As Long
58    Dim lngFiscalMonthCol As Long
59    Dim lngHoursCol As Long
60    Dim lngRateSets As Long
61    Dim lngCol As Long
62    Dim lngOriginalRateSet As Long
63    Dim lngTasks As Long
64    Dim lngTask As Long
65    Dim lngWeekCol As Long
66    Dim lngExport As Long
67    Dim lngField As Long
68    Dim lngRateSet As Long
69    Dim lngRow As Long
70    'variants
71    Dim vParts As Variant
72    Dim aResult() As Variant
73    Dim vKey As Variant
74    Dim vData As Variant
75    Dim vChk As Variant
76    Dim vRateSet As Variant
77    Dim aUserFields() As Variant
78    Dim vFiscalCalendar As Variant
79    'booleans
80    Dim blnErrorTrapping As Boolean
81    Dim blnFiscal As Boolean
82    Dim blnExportAssociatedBaseline As Boolean
83    Dim blnExportFullBaseline As Boolean
84    Dim blnIncludeCosts As Boolean
85
86    blnErrorTrapping = cptErrorTrapping
87    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
88
89    Application.StatusBar = "Confirming Status Date..."
90    myResourceDemand_frm.lblStatus.Caption = "Confirming Status Date..."
91
92    If IsDate(ActiveProject.StatusDate) Then
93      dtStart = ActiveProject.StatusDate
94      If ActiveProject.ProjectStart > dtStart Then dtStart = ActiveProject.ProjectStart
95    Else
96      Application.StatusBar = "Please enter a Status Date."
97      MsgBox "Please enter a Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
98      Application.StatusBar = ""
99      GoTo exit_here
100   End If
101
102   'save settings, build header
103   strHeader = "PROJECT,"
104   With myResourceDemand_frm
105     Application.StatusBar = "Saving user settings..."
106     aUserFields = .lboExport.List()
107     For lngExport = 0 To UBound(aUserFields, 1)
108       lngField = aUserFields(lngExport, 0)
109       strCFN = CustomFieldGetName(lngField)
110       If Len(strCFN) > 0 Then
111         strHeader = strHeader & UCase(strCFN) & ","
112       Else
113         strHeader = strHeader & UCase(FieldConstantToFieldName(lngField)) & ","
114       End If
115     Next lngExport
116     strHeader = strHeader & "[UID] TASK,RESOURCE_NAME,CLASS,"
117     .lblStatus.Caption = "Saving user settings..."
118     cptSaveSetting "ResourceDemand", "cboMonths", .cboMonths.Value
119     blnFiscal = .cboMonths.Value = 1
120     cptSaveSetting "ResourceDemand", "cboWeeks", .cboWeeks.Value
121     cptSaveSetting "ResourceDemand", "cboWeekday", .cboWeekday.Value
122     cptSaveSetting "ResourceDemand", "chkCosts", IIf(.chkCosts, 1, 0)
123     blnIncludeCosts = .chkCosts
124     If blnIncludeCosts Then
125       lngItem = 0
126       For Each vChk In Split("A,B,C,D,E", ",")
127         strRateSets = strRateSets & IIf(.Controls("chk" & vChk), lngItem & ",", "")
128         lngItem = lngItem + 1
129       Next
130       If Len(strRateSets) > 0 Then strRateSets = Left(strRateSets, Len(strRateSets) - 1)
131       lngRateSets = UBound(Split(strRateSets, ",")) + 1
132       cptSaveSetting "ResourceDemand", "CostSets", strRateSets
133       strHeader = strHeader & "RATE_TABLE,ACTIVE,"
134     End If
135     If blnFiscal Then
136       strHeader = strHeader & "FISCAL_MONTH,"
137     Else
138       strHeader = strHeader & "WEEK,MONTH,"
139     End If
140     strHeader = strHeader & "HOURS"
141     If blnIncludeCosts Then
142       strHeader = strHeader & ",COST"
143     End If
144     cptDeleteSetting "ResourceDemand", "chkBaseline"
145     blnExportAssociatedBaseline = .chkAssociatedBaseline = True
146     cptSaveSetting "ResourceDemand", "chkAssociatedBaseline", IIf(blnExportAssociatedBaseline, 1, 0)
147     blnExportFullBaseline = .chkFullBaseline = True
148     cptSaveSetting "ResourceDemand", "chkFullBaseline", IIf(blnExportFullBaseline, 1, 0)
149     cptDeleteSetting "ResourceDemand", "chkNonLabor"
150   End With
151
152   strFileName = cptDir & "\settings\cpt-export-resource-userfields.adtg."
153   Set oSettings = CreateObject("ADODB.Recordset")
154   With oSettings
155     .Fields.Append "Field Constant", adVarChar, 255
156     .Fields.Append "Custom Field Name", adVarChar, 255
157     .Open
158     strSettings = "Week=" & myResourceDemand_frm.cboWeeks & ";"
159     strSettings = strSettings & "Weekday=" & myResourceDemand_frm.cboWeekday & ";"
160     strSettings = strSettings & "Costs=" & myResourceDemand_frm.chkCosts & ";"
161     strSettings = strSettings & "AssociatedBaseline=" & blnExportAssociatedBaseline & ";"
162     strSettings = strSettings & "FullBaseline=" & blnExportFullBaseline & ";"
163     strSettings = strSettings & "RateSets="
164     For Each vChk In Split("A,B,C,D,E", ",")
165       strFields = strFields & IIf(myResourceDemand_frm.Controls("chk" & vChk), vChk & ",", "")
166     Next vChk
167     .AddNew Array(0, 1), Array("settings", strSettings)
168     .Update
169     'save userfields
170     For lngExport = 0 To UBound(aUserFields, 1)
171       .AddNew Array(0, 1), Array(aUserFields(lngExport, 0), aUserFields(lngExport, 1))
172       .Update
173     Next lngExport
174     If Dir(strFileName) <> vbNullString Then Kill strFileName
175     .Save strFileName, adPersistADTG
176     .Close
177   End With
178
179   Application.StatusBar = "Preparing to export..."
180   myResourceDemand_frm.lblStatus.Caption = "Preparing to export..."
181
182   If ActiveProject.Subprojects.Count = 0 Then
183     lngTasks = ActiveProject.Tasks.Count
184   Else
185     cptSpeed True
186     strView = ActiveWindow.TopPane.View.Name
187     ViewApply "Gantt Chart"
188     FilterClear
189     GroupClear
190     SelectAll
191     OptionsViewEx DisplaySummaryTasks:=True
192     On Error Resume Next
193     If Not OutlineShowAllTasks Then
194       Sort "ID", , , , , , False, True
195       OutlineShowAllTasks
196     End If
197     If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
198     SelectAll
199     lngTasks = ActiveSelection.Tasks.Count
200     ViewApply strView
201     cptSpeed False
202   End If
203
204   If blnFiscal Then 'get the fiscal calendar
205     Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
206     ReDim vFiscalCalendar(0 To 1, 0 To oCalendar.Exceptions.Count)
207     For Each oException In oCalendar.Exceptions
208       vFiscalCalendar(0, oException.Index) = oException.Start
209       vFiscalCalendar(1, oException.Index) = oException.Name
210     Next oException
211   End If
212
213   'Key=PROJECT|{USER_FIELD}|[UID] TASK|RESOURCE_NAME|CLASS|COST_SET|ACTIVE|MONTH
214   'Value=HOURS|COST
215
216   'iterate over tasks
217   Application.StatusBar = "Getting Excel..."
218   myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
219   'set reference to Excel
220 '  On Error Resume Next
221 '  Set oExcel = GetObject(, "Excel.Application")
222 '  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
223 '  If oExcel Is Nothing Then
224     Set oExcel = CreateObject("Excel.Application")
225 '  End If
226
227   Set oEstimates = CreateObject("Scripting.Dictionary")
228   For Each oTask In ActiveProject.Tasks
229     If oTask Is Nothing Then GoTo next_task 'skip blank lines
230     If oTask.ExternalTask Then GoTo next_task 'skip external tasks
231     If oTask.Summary Then GoTo next_task 'skip summary task
232     If Not oTask.Active Then GoTo next_task 'skip inactive tasks
233     If Not blnExportAssociatedBaseline And Not blnExportFullBaseline Then
234       If oTask.RemainingDuration = 0 Then GoTo next_task
235     End If
236
237     'capture oTask data common to all oAssignments
238     strTask = oTask.Project
239
240     'get custom field values
241     For lngExport = 0 To UBound(aUserFields, 1) 'myResourceDemand_frm.lboExport.ListCount - 1
242       lngField = aUserFields(lngExport, 0)
243       strTask = strTask & "|" & Trim(Replace(oTask.GetField(lngField), "|", "-"))
244     Next lngExport
245
246     strTask = strTask & "|[" & oTask.UniqueID & "] " & Replace(Replace(oTask.Name, "|", "-"), Chr(34), Chr(39))
247
248     'examine every oAssignment on the task
249     For Each oAssignment In oTask.Assignments
250
251       'capture original rate set
252       lngOriginalRateSet = oAssignment.CostRateTable
253
254       'skip non-labor entirely
255       If oAssignment.ResourceType <> pjResourceTypeWork Then GoTo next_assignment 'skip non-labor entirely
256
257       'skip completed tasks for ETC
258       If IsDate(oTask.ActualFinish) Then GoTo export_baseline 'NOT Exit For
259
260
261       'capture remaining work (ETC)
262       If IsDate(oTask.Stop) Then 'capture the unstatused / remaining portion
263         dtStart = oTask.Resume
264       Else 'capture the entire unstarted task
265         dtStart = oTask.Start
266       End If
267       dtFinish = oTask.Finish
268
269       If blnFiscal Then
270         'Set oTSVS_WORK = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleDays, 1)
271         Set oTSVS_WORK = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
272       Else
273         Set oTSVS_WORK = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
274       End If
275
276       For Each oTSV In oTSVS_WORK
277         If Val(oTSV.Value) = 0 Then GoTo next_tsv_etc
278         'capture common oAssignment data
279         strKey = strTask & "|" & oAssignment.ResourceName & "|ETC" 'keep this here
280         dtStartDate = oTSV.StartDate
281         dtEndDate = oTSV.EndDate
282         'capture (and subtract) actual work, leaving ETC/Remaining Work
283         If blnFiscal Then
284           'Set oTSVS_AW = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualWork, pjTimescaleDays, 1)
285           Set oTSVS_AW = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
286         Else
287           Set oTSVS_AW = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
288         End If
289         dblWork = (Val(oTSV.Value) - Val(oTSVS_AW(1))) / 60
290         If dblWork = 0 Then GoTo next_tsv_etc
291
292         If blnIncludeCosts Then
293           strKey = strKey & "|" & Choose(oAssignment.CostRateTable + 1, "A", "B", "C", "D", "E")
294           strKey = strKey & "|TRUE"
295         End If
296
297         If blnFiscal Then
298           strKey = strKey & "|" & cptGetFiscalMonthOfDay(dtStartDate, vFiscalCalendar)
299         Else
300           'apply user settings for week identification
301           With myResourceDemand_frm
302             If .cboWeeks = "Beginning" Then
303               If .cboWeekday = "Monday" Then
304                 dtWeek = DateAdd("d", 2 - Weekday(dtStartDate), dtStartDate)
305               End If
306             ElseIf .cboWeeks = "Ending" Then
307               If .cboWeekday = "Friday" Then
308                 dtWeek = DateAdd("d", 6 - Weekday(dtStartDate), dtStartDate)
309               ElseIf .cboWeekday = "Saturday" Then
310                 dtWeek = DateAdd("d", 7 - Weekday(dtStartDate), dtStartDate)
311               End If
312             End If
313           End With
314           strKey = strKey & "|" & dtWeek & "|" & Format(dtWeek, "yyyymm")
315         End If
316
317         'add work without cost yet
318         If oEstimates.Exists(strKey) Then
319           If blnIncludeCosts Then
320             dblWork = dblWork + Split(oEstimates(strKey), "|")(0) 'add
321             dblCost = Split(oEstimates(strKey), "|")(1) 'keep
322             oEstimates(strKey) = dblWork & "|" & dblCost
323           Else
324             oEstimates(strKey) = oEstimates(strKey) + dblWork
325           End If
326         Else
327           If blnIncludeCosts Then
328             oEstimates.Add strKey, dblWork & "|" & 0 'dblCost
329           Else
330             oEstimates.Add strKey, dblWork
331           End If
332         End If
333
334         'get default costs
335         If blnIncludeCosts Then
336           'get active cost
337           If blnFiscal Then
338             'Set oTSVS_COST = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledCost, pjTimescaleDays, 1)
339             Set oTSVS_COST = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
340             'get actual cost
341             'Set oTSVS_AC = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualCost, pjTimescaleDays, 1)
342             Set oTSVS_AC = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
343           Else
344             Set oTSVS_COST = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
345             'get actual cost
346             Set oTSVS_AC = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
347           End If
348           'subtract actual cost from cost to get remaining cost
349           dblCost = Val(oTSVS_COST(1).Value) - Val(oTSVS_AC(1))
350
351           'add cost without work
352           If oEstimates.Exists(strKey) Then
353             If blnIncludeCosts Then
354               dblWork = Split(oEstimates(strKey), "|")(0) 'keep
355               dblCost = dblCost + Split(oEstimates(strKey), "|")(1) 'add
356               oEstimates(strKey) = dblWork & "|" & dblCost
357             'Else
358               'oEstimates(strKey) = oEstimates(strKey) + dblWork
359             End If
360           Else
361             If blnIncludeCosts Then
362               'Stop 'uh oh
363               oEstimates.Add strKey, 0 & "|" & dblCost 'this should never happen
364             'Else
365               'oEstimates.Add strKey, dblWork
366             End If
367           End If
368         End If
369
next_tsv_etc:
371       Next oTSV
372
373       If lngRateSets > 0 Then
374         'silly to have to repeat it, but changing cost rate tables is expensive
375         'better to do it once per rate table, per assignment
376         'than to do it once per rate table, per assignment, per timescalevalue
377         For Each vRateSet In Split(strRateSets, ",")
378           If CLng(vRateSet) = lngOriginalRateSet Then GoTo next_rate_set
379
380           For Each oTSV In oTSVS_WORK
381             'capture common oAssignment data
382             strKey = strTask & "|" & oAssignment.ResourceName & "|ETC" 'keep this here
383             'capture (and subtract) actual work, leaving ETC/Remaining Work
384             If blnFiscal Then
385               'Set oTSVS_AW = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualWork, pjTimescaleDays, 1)
386               Set oTSVS_AW = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
387             Else
388               Set oTSVS_AW = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
389             End If
390             dblWork = (Val(oTSV.Value) - Val(oTSVS_AW(1))) / 60
391             If dblWork = 0 Then GoTo next_tsv_rs
392
393             If blnIncludeCosts Then
394               strKey = strKey & "|" & Choose(CLng(vRateSet) + 1, "A", "B", "C", "D", "E")
395               strKey = strKey & "|FALSE"
396             End If
397
398             If blnFiscal Then
399               strKey = strKey & "|" & cptGetFiscalMonthOfDay(dtStartDate, vFiscalCalendar)
400             Else
401               'apply user settings for week identification
402               With myResourceDemand_frm
403                 If .cboWeeks = "Beginning" Then
404                   If .cboWeekday = "Monday" Then
405                     dtWeek = DateAdd("d", 2 - Weekday(dtStartDate), dtStartDate)
406                   End If
407                 ElseIf .cboWeeks = "Ending" Then
408                   If .cboWeekday = "Friday" Then
409                     dtWeek = DateAdd("d", 6 - Weekday(dtStartDate), dtStartDate)
410                   ElseIf .cboWeekday = "Saturday" Then
411                     dtWeek = DateAdd("d", 7 - Weekday(dtStartDate), dtStartDate)
412                   End If
413                 End If
414               End With
415               strKey = strKey & "|" & dtWeek & "|" & Format(dtWeek, "yyyymm")
416             End If
417
418             'add work without cost yet
419             If oEstimates.Exists(strKey) Then
420               If blnIncludeCosts Then
421                 dblWork = dblWork + Split(oEstimates(strKey), "|")(0) 'add
422                 dblCost = Split(oEstimates(strKey), "|")(1) 'keep
423                 oEstimates(strKey) = dblWork & "|" & dblCost
424               Else
425                 oEstimates(strKey) = oEstimates(strKey) + dblWork
426               End If
427             Else
428               If blnIncludeCosts Then
429                 oEstimates.Add strKey, dblWork & "|" & 0 'dblCost
430               Else
431                 oEstimates.Add strKey, dblWork
432               End If
433             End If
434
435             'get active cost
436             If oAssignment.CostRateTable <> CLng(vRateSet) Then oAssignment.CostRateTable = CLng(vRateSet) 'very expensive
437             If blnFiscal Then
438               'Set oTSVS_COST = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledCost, pjTimescaleDays, 1)
439               Set oTSVS_COST = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
440               'get actual cost
441               'Set oTSVS_AC = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualCost, pjTimescaleDays, 1)
442               Set oTSVS_AC = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
443             Else
444               Set oTSVS_COST = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
445               'get actual cost
446               Set oTSVS_AC = oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
447             End If
448             'subtract actual cost from cost to get remaining cost
449             dblCost = Val(oTSVS_COST(1).Value) - Val(oTSVS_AC(1))
450
451             'add cost without work
452             If oEstimates.Exists(strKey) Then
453               If blnIncludeCosts Then
454                 dblWork = Split(oEstimates(strKey), "|")(0) 'keep
455                 dblCost = dblCost + Split(oEstimates(strKey), "|")(1) 'add
456                 oEstimates(strKey) = dblWork & "|" & dblCost
457               End If
458             Else
459               'this should never happen
460             End If
461
next_tsv_rs:
463           Next oTSV
next_rate_set:
465         Next vRateSet
466         If oAssignment.CostRateTable <> lngOriginalRateSet Then oAssignment.CostRateTable = lngOriginalRateSet
467       End If
468
export_baseline:
470       If blnExportAssociatedBaseline Or blnExportFullBaseline Then
471         dtStart = oExcel.WorksheetFunction.Min(oTask.Start, IIf(oTask.BaselineStart = "NA", oTask.Start, oTask.BaselineStart)) 'works with forecast, actual, and baseline start
472         dtFinish = oExcel.WorksheetFunction.Max(oTask.Finish, IIf(oTask.BaselineFinish = "NA", oTask.Finish, oTask.BaselineFinish)) 'works with forecast, actual, and baseline finish
473         'Set oTSVS_BCWS = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleDays, 1)
474         Set oTSVS_BCWS = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks, 1)
475         For Each oTSV In oTSVS_BCWS
476           If Val(oTSV.Value) = 0 Then GoTo next_tsv_bcws
477           strKey = strTask & "|" & oAssignment.ResourceName & "|BCWS" 'keep this here
478           dtStartDate = oTSV.StartDate
479           dtEndDate = oTSV.EndDate
480           dblWork = Val(oTSV.Value) / 60
481           If blnIncludeCosts Then
482             strKey = strKey & "|BASELINED|TRUE"
483             'dblCost = Val(oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledBaselineCost, pjTimescaleDays, 1)(1).Value)
484             dblCost = Val(oAssignment.TimeScaleData(dtStartDate, dtEndDate, pjAssignmentTimescaledBaselineCost, pjTimescaleWeeks, 1)(1).Value)
485           End If
486           If blnFiscal Then
487             'get fiscal month of day
488             strKey = strKey & "|" & cptGetFiscalMonthOfDay(dtStartDate, vFiscalCalendar)
489           Else
490             'apply user settings for week identification
491             With myResourceDemand_frm
492               If .cboWeeks = "Beginning" Then
493                 If .cboWeekday = "Monday" Then
494                   dtWeek = DateAdd("d", 2 - Weekday(dtStartDate), dtStartDate)
495                 End If
496               ElseIf .cboWeeks = "Ending" Then
497                 If .cboWeekday = "Friday" Then
498                   dtWeek = DateAdd("d", 6 - Weekday(dtStartDate), dtStartDate)
499                 ElseIf .cboWeekday = "Saturday" Then
500                   dtWeek = DateAdd("d", 7 - Weekday(dtStartDate), dtStartDate)
501                 End If
502               End If
503             End With
504             strKey = strKey & "|" & dtWeek & "|" & Format(dtWeek, "yyyymm")
505           End If
506           If oEstimates.Exists(strKey) Then
507             If blnIncludeCosts Then
508               dblWork = dblWork + Split(oEstimates(strKey), "|")(0)
509               dblCost = dblCost + Split(oEstimates(strKey), "|")(1)
510               oEstimates(strKey) = dblWork & "|" & dblCost
511             Else
512               dblWork = dblWork + Split(oEstimates(strKey), "|")(0)
513               oEstimates(strKey) = dblWork
514             End If
515           Else
516             If blnIncludeCosts Then
517               oEstimates.Add strKey, dblWork & "|" & dblCost
518             Else
519               oEstimates.Add strKey, dblWork
520             End If
521           End If
next_tsv_bcws:
523         Next oTSV
524       End If
next_assignment:
526       'restore original rate set
527       If oAssignment.CostRateTable <> lngOriginalRateSet Then oAssignment.CostRateTable = lngOriginalRateSet
528     Next oAssignment
next_task:
530     lngTask = lngTask + 1
531     Application.StatusBar = "Exporting " & Format(lngTask, "#,##0") & " of " & Format(lngTasks, "#,##0") & "...(" & Format(lngTask / lngTasks, "0%") & ")"
532     myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
533     myResourceDemand_frm.lblProgress.Width = (lngTask / lngTasks) * myResourceDemand_frm.lblStatus.Width
534     DoEvents
535   Next oTask
536
537   If oEstimates.Count > 0 Then
538     Application.StatusBar = "Creating Workbook..."
539     myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
540     Set oWorkbook = oExcel.Workbooks.Add
541     Set oWorksheet = oWorkbook.Sheets(1)
542     'header
543     oWorksheet.[A1].Resize(1, UBound(Split(strHeader, ",")) + 1) = Split(strHeader, ",")
544     'data
545     ReDim aResult(1 To oEstimates.Count, 1 To UBound(Split(strHeader, ",")) + 1)
546     oWorksheet.[A1].AutoFilter
547     lngRow = 1
548     lngCols = UBound(Split(strHeader, ",")) + 1
549     For Each vKey In oEstimates.Keys
550       vParts = Split(vKey, "|")
551       For lngCol = 1 To (UBound(vParts, 1) + 1)
552         aResult(lngRow, lngCol) = vParts(lngCol - 1)
553       Next lngCol
554       If blnIncludeCosts Then
555         aResult(lngRow, lngCols - 1) = Split(oEstimates(vKey), "|")(0)
556         aResult(lngRow, lngCols) = Split(oEstimates(vKey), "|")(1)
557       Else
558         aResult(lngRow, lngCols) = oEstimates(vKey)
559       End If
560       lngRow = lngRow + 1
561     Next vKey
562     oWorksheet.[A2].Resize(UBound(aResult, 1), UBound(aResult, 2)).Value = aResult
563   End If
564
565 '  'is previous run still open?
566 '  On Error Resume Next
567 '  strFileName = Environ("TEMP") & "\ExportResourceDemand.xlsx"
568 '  Set oWorkbook = oExcel.oWorkbooks(strFileName)
569 '  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
570 '  If Not oWorkbook Is Nothing Then oWorkbook.Close False
571 '  On Error Resume Next
572 '  Set oWorkbook = oExcel.Workbooks(Environ("TEMP") & "\ExportResourceDemand.xlsx")
573 '  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
574 '  If Not oWorkbook Is Nothing Then 'add timestamp to existing file
575 '    If oWorkbook.Application.Visible = False Then oWorkbook.Application.Visible = True
576 '    strMsg = "'" & strFileName & "' already exists and is open."
577 '    strFileName = Replace(strFileName, ".xlsx", "_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".xlsx")
578 '    strMsg = strMsg & "Your new file will be saved as:" & vbCrLf & strFileName
579 '    MsgBox strMsg, vbExclamation + vbOKOnly, "File Exists and is Open"
580 '  End If
581
582   Application.StatusBar = "Saving workbook..."
583   myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
584
585   On Error Resume Next
586   If oWorkbook Is Nothing Then GoTo exit_here 'todo
587   If Dir(Environ("TEMP") & "\ExportResourceDemand.xlsx") <> vbNullString Then Kill Environ("TEMP") & "\ExportResourceDemand.xlsx"
588   If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
589   MsgBox "If your company requires security classifications, please make them from within the Excel Window.", vbExclamation + vbOKOnly, "Heads up"
590   oExcel.Visible = True
591   oExcel.WindowState = xlNormal
592   If Dir(Environ("TEMP") & "\ExportResourceDemand.xlsx") <> vbNullString Then 'kill failed, rename it
593     oWorkbook.SaveAs Environ("TEMP") & "\ExportResourceDemand_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".xlsx", 51
594   Else
595     oWorkbook.SaveAs Environ("TEMP") & "\ExportResourceDemand.xlsx", 51
596   End If
597   oExcel.Visible = False
598
599   If blnFiscal Then
600     Application.StatusBar = "Extracting Fiscal Periods..."
601     myResourceDemand_frm.lblStatus.Caption = "Extracting Fiscal Periods..."
602     Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
603     With oWorksheet
604       .Name = "FiscalPeriods"
605       .[A1:B1] = Array("fisc_end", "label")
606       .[A2].Resize(UBound(vFiscalCalendar, 2), UBound(vFiscalCalendar, 1) + 1).Value = oExcel.WorksheetFunction.Transpose(vFiscalCalendar)
607       Set oListObject = .ListObjects.Add(xlSrcRange, .Range(.[A1].End(xlToRight), .[A1].End(xlDown)), , xlYes)
608       oListObject.Name = "FISCAL"
609       'add Holidays table
610       .[E1] = "EXCEPTIONS"
611     End With
612     'just...go get the exceptions
613     Set oCalendar = ActiveProject.Calendar
614     If oCalendar.Exceptions.Count > 0 Then
615       Set oWorksheet = oWorkbook.Worksheets.Add(After:=oWorkbook.Worksheets(oWorkbook.Worksheets.Count))
616       oWorksheet.Name = "Exceptions"
617       Set oWorksheet = oWorkbook.Worksheets.Add(After:=oWorkbook.Worksheets(oWorkbook.Worksheets.Count))
618       oWorksheet.Name = "WorkWeeks"
619       cptExportCalendarExceptions oWorkbook, oCalendar, True
620       With oWorksheet
621         .Activate
622         oExcel.ActiveWindow.Zoom = 85
623         .Columns.AutoFit
624       End With
625       Set oWorksheet = oWorkbook.Worksheets("Exceptions")
626       With oWorksheet
627         .Activate
628         oExcel.ActiveWindow.Zoom = 85
629         .Columns.AutoFit
630         .Outline.ShowLevels Rowlevels:=1
631       End With
632       Set oWorksheet = oWorkbook.Worksheets("FiscalPeriods")
633       With oWorksheet
634         .Activate
635         .[E2].Formula2 = "=UNIQUE(Exceptions!" & oWorkbook.Sheets("Exceptions").Range(oWorkbook.Sheets("Exceptions").[C2], oWorkbook.Sheets("Exceptions").[C2].End(xlDown)).Address & ")"
636         .Range(.[E2], .[E2].End(xlDown)).NumberFormat = "m/d/YYYY"
637         vData = .Range(.[E2], .[E2].End(xlDown))
638         .Range(.[E2], .[E2].End(xlDown)) = vData
639         'convert to a table
640         Set oListObject = .ListObjects.Add(xlSrcRange, .Range(.[E1], .[E2].End(xlDown)), , xlYes)
641         'reset oCalendar
642         Set oCalendar = ActiveProject.Calendar
643         .Columns(6).ColumnWidth = 1
644         .[G3] = "Fiscal periods imported from 'cptFiscalCalendar'"
645         With .[G3:L3]
646           .Merge
647           .HorizontalAlignment = xlCenter
648           .Style = "Note"
649         End With
650         .[G4] = "Exceptions imported from '" & oCalendar.Name & "'"
651         With .[G4:L4]
652           .Merge
653           .HorizontalAlignment = xlCenter
654           .Style = "Note"
655         End With
656       End With
657     Else
658       'convert to a table
659       Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[E1], oWorksheet.[E2]), , xlYes)
660     End If
661     oExcel.ActiveWindow.DisplayGridlines = False
662     oExcel.ActiveWindow.Zoom = 85
663     oListObject.Name = "EXCEPTIONS"
664     'add efficiency factor entry
665     With oWorksheet.[G1]
666       .Value = "Efficiency:"
667       .EntireColumn.AutoFit
668     End With
669     With oWorksheet
670       With .[H1]
671         .Value = 1
672         .Style = "Percent"
673         .Style = "Input"
674       End With
675       .Names.Add "efficiency_factor", .[H1]
676     End With
677     'add HPM formula
678     Application.StatusBar = "Calculating HPM..."
679     myResourceDemand_frm.lblStatus.Caption = "Calculating HPM..."
680     oWorksheet.[C1].Value = "hpm"
681     oWorksheet.[C3].Formula = "=IFERROR(NETWORKDAYS(A2+1,[@[fisc_end]],EXCEPTIONS)*(8*efficiency_factor),0)"
682   End If
683
684   Set oWorksheet = oWorkbook.Sheets(1)
685   With oWorksheet
686     .Name = "SourceData"
687     lngHoursCol = .Rows(1).Find("HOURS", lookat:=1).Column '1=xlWhole
688     If Not blnFiscal Then
689       lngWeekCol = oWorksheet.Rows(1).Find("WEEK", lookat:=1).Column '1=xlWhole
690     End If
691     lngLastCol = .[A2].End(xlToRight).Column
692     'number formats
693     For lngCol = 1 To lngLastCol
694       If cptRxTest(.Cells(1, lngCol), "(HOURS|COST|FTE)") Then
695         .Columns(lngCol).NumberFormat = "_(* #,##0.00000000_);_(* (#,##0.00000000);_(* ""-""??_);_(@_)"
696       End If
697     Next lngCol
698     'add note on CostRateTable column
699     If blnIncludeCosts Then
700       lngCol = .Rows(1).Find("RATE_TABLE", lookat:=1).Column
701       .Cells(1, lngCol).AddComment "Rate Table Applied in the Project"
702     End If
703     'add fte for non-fiscal
704     If Not blnFiscal Then
705       'create FTE_WEEK column
706       Set oRange = .[A1].End(xlToRight).End(xlDown).Offset(0, 1)
707       Set oRange = .Range(oRange, .[A1].End(xlToRight).Offset(1, 1))
708       If blnFiscal Then 'fiscal
709         'get fiscal_month column
710         lngFiscalMonthCol = .Rows(1).Find(what:="FISCAL_MONTH", lookat:=xlWhole).Column
711         oRange.FormulaR1C1 = "=RC" & lngHoursCol & "/NETWORKDAYS(RC" & lngWeekCol & "-7,RC" & lngWeekCol & ",EXCEPTIONS)"
712       Else
713         oRange.FormulaR1C1 = "=RC" & lngHoursCol & "/40"
714       End If
715       .[A1].End(xlToRight).Offset(0, 1).Value = "FTE_WEEK"
716     End If
717     'create FTE/FTE_MONTH column
718     Set oRange = .[A1].End(xlToRight).Offset(1, 1)
719     Set oRange = .Range(oRange, .Cells(.UsedRange.Rows.Count, oRange.Column))
720     lngHoursCol = .Rows(1).Find("HOURS", lookat:=1).Column '1=xlWhole
721     If blnFiscal Then
722       lngFiscalMonthCol = .Rows(1).Find("FISCAL_MONTH", lookat:=1).Column '1=xlWhole
723       oRange.FormulaR1C1 = "=RC" & lngHoursCol & "/LOOKUP(RC" & lngFiscalMonthCol & ",FISCAL[label],FISCAL[hpm])"
724       .[A1].End(xlToRight).Offset(0, 1).Value = "FTE"
725     Else
726       lngWeekCol = .Rows(1).Find("WEEK", lookat:=1).Column
727       oRange.FormulaR1C1 = "=RC" & lngHoursCol & "/160" 'todo: can we do something smarter?
728       .[A1].End(xlToRight).Offset(0, 1).Value = "FTE_MONTH"
729     End If
730     'capture the range of data to feed as variable to PivotTable
731     Set oRange = .Range(.[A1].End(xlDown), .[A1].End(xlToRight))
732     strRange = .Name & "!" & Replace(oRange.Address, "$", "")
733   End With 'SourceData Worksheet
734
735   'add a new Worksheet for the oPivotTable
736   Set oWorksheet = oWorkbook.Sheets.Add(Before:=oWorkbook.Sheets("SourceData"))
737   'rename the new Worksheet
738   oWorksheet.Name = "ResourceDemand"
739
740   Application.StatusBar = "Creating PivotTable..."
741   myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
742
743   'create the PivotTable
744   oWorkbook.PivotCaches.Create(SourceType:=1, _
        SourceData:=strRange, Version:= _
        3).CreatePivotTable TableDestination:="ResourceDemand!R3C1", TableName:="RESOURCE_DEMAND", DefaultVersion:=3
747   Set oPivotTable = oWorksheet.PivotTables(1)
748   If blnFiscal Then
749     oPivotTable.AddFields Array("RESOURCE_NAME", "[UID] TASK"), Array("FISCAL_MONTH")
750     oPivotTable.AddDataField oPivotTable.PivotFields("FTE"), "FTE ", -4157
751   Else
752     If ActiveProject.Subprojects.Count > 0 Then
753       oPivotTable.AddFields Array("RESOURCE_NAME", "PROJECT", "[UID] TASK"), Array("WEEK")
754     Else
755       oPivotTable.AddFields Array("RESOURCE_NAME", "[UID] TASK"), Array("WEEK")
756     End If
757     oPivotTable.AddDataField oPivotTable.PivotFields("FTE_WEEK"), "FTE_WEEK ", -4157
758   End If
759
760   'set default to ETC
761   If blnExportAssociatedBaseline Or blnExportFullBaseline Then
762     With oPivotTable
763       With .PivotFields("CLASS")
764         .Orientation = xlPageField
765         .Position = 1
766         .ClearAllFilters
767         .CurrentPage = "ETC"
768       End With
769       If lngRateSets > 0 Then
770         With .PivotFields("ACTIVE")
771           .Orientation = xlPageField
772           .Position = 1
773           .ClearAllFilters
774           .CurrentPage = "TRUE"
775         End With
776       End If
777     End With
778   End If
779
780   'format the oPivotTable
781   With oPivotTable
782     .ShowDrillIndicators = True
783     .EnableDrilldown = True
784     .PivotCache.MissingItemsLimit = xlMissingItemsNone
785     .PivotFields("RESOURCE_NAME").ShowDetail = False
786     .TableStyle2 = "PivotStyleLight16"
787     .PivotSelect "", 2, True
788   End With
789   oExcel.Selection.Style = "Comma"
790   With oExcel.Selection
791     With .FormatConditions
792       .Delete
793       .AddColorScale ColorScaleType:=2
794     End With
795     With .FormatConditions(1)
796       .SetFirstPriority
797       With .ColorScaleCriteria(1)
798         .Type = 1 '1=xlConditionValueLowestValue
799         .FormatColor.Color = 10285055
800         .FormatColor.TintAndShade = 0
801       End With
802       With .ColorScaleCriteria(2)
803         .Type = 2 '2=xlConditionValueHighestValue
804         .FormatColor.Color = 2650623
805         .FormatColor.TintAndShade = 0
806       End With
807       .ScopeType = 1 '1=xlFieldsScope
808     End With
809   End With
810
811   Application.StatusBar = "Building header..."
812   myResourceDemand_frm.lblStatus = Application.StatusBar
813
814   'add a title
815   With oWorksheet
816     .Rows("1:3").EntireRow.Insert
817     .[A2] = "Status Date: " & FormatDateTime(ActiveProject.StatusDate, vbShortDate)
818     .[A2].EntireColumn.AutoFit
819     With .[A1]
820       .Value = "REMAINING WORK IN IMS: " & cptRegEx(ActiveProject.Name, "[^\\/]{1,}$")
821       With .Font
822         .Bold = True
823         .Italic = True
824         .Size = 14
825       End With
826     End With
827     .[A1:F1].Merge
828     'revise according to user options
829     If blnFiscal Then
830       .[B2] = "FTE by Fiscal Month"
831     Else
832       .[B2] = "FTE by Weeks " & myResourceDemand_frm.cboWeeks.Value & " " & myResourceDemand_frm.cboWeekday.Value
833     End If
834     oPivotTable.DataBodyRange.Select
835     oExcel.ActiveWindow.FreezePanes = True
836     .[A2].Select
837     'make it nice
838     oExcel.ActiveWindow.Zoom = 85
839   End With
840
841   Application.StatusBar = "Creating PivotChart..."
842   myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
843
844   'create a PivotChart
845   Set oWorksheet = oWorkbook.Sheets("SourceData")
846   With oWorksheet
847     .Activate
848     .[A2].Select
849     .[A2].EntireColumn.AutoFit
850     oExcel.ActiveWindow.Zoom = 85
851     oExcel.ActiveWindow.FreezePanes = True
852     .Cells.EntireColumn.AutoFit
853   End With
854   Set oWorksheet = oWorkbook.Sheets.Add
855   With oWorksheet
856     .Name = "PivotChart_Source"
857     Set oPivotTable = oWorkbook.Worksheets("ResourceDemand").PivotTables("RESOURCE_DEMAND")
858     oPivotTable.PivotCache.CreatePivotTable TableDestination:="PivotChart_Source!R1C1", TableName:="PivotTable1", DefaultVersion:=3
859     .Activate
860     .[A1].Select
861     Set oChart = .Shapes.AddChart2.Chart
862     Set oRange = .Range(.[A1].End(-4161), .[A1].End(-4121))
863   End With
864   oChart.SetSourceData Source:=oRange
865   oWorkbook.ShowPivotChartActiveFields = True
866   oChart.ChartType = 76 'xlAreaStacked
867   Set oPivotChartTable = oChart.PivotLayout.PivotTable
868   If blnFiscal Then
869     With oPivotChartTable.PivotFields("FISCAL_MONTH")
870       .Orientation = 1 'xlRowField
871       .Position = 1
872     End With
873   Else
874     With oPivotChartTable.PivotFields("WEEK")
875       .Orientation = 1 'xlRowField
876       .Position = 1
877     End With
878   End If
879   oPivotChartTable.AddDataField oPivotChartTable.PivotFields("HOURS"), "Sum of HOURS", -4157
880   With oPivotChartTable.PivotFields("RESOURCE_NAME")
881     .Orientation = 2 'xlColumnField
882     .Position = 1
883   End With
884   If blnExportAssociatedBaseline Or blnExportFullBaseline Then
885     'set default to ETC
886     With oPivotChartTable.PivotFields("CLASS")
887       .Orientation = xlPageField
888       .Position = 1
889       .ClearAllFilters
890       .CurrentPage = "ETC"
891     End With
892   Else
893     If Not blnFiscal Then
894       oPivotTable.PivotFields("WEEK").PivotFilters.Add Type:=33, Value1:=ActiveProject.StatusDate '33 = xlAfter
895     End If
896   End If
897   With oChart
898     .ClearToMatchStyle
899     .ChartStyle = 34
900     .ClearToMatchStyle
901     .SetElement (msoElementChartTitleAboveChart)
902     .ChartTitle.Text = "Resource Demand"
903     .Location 1, "PivotChart" 'xlLocationAsNewSheet = 1
904   End With
905   Set oWorksheet = oWorkbook.Sheets("PivotChart_Source")
906   oWorksheet.Visible = False
907
908   'add legend
909   oExcel.ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
910   oExcel.ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "HOURS"
911
912   'export selected cost rate tables to oWorksheet
913   If blnIncludeCosts Then
914     Application.StatusBar = "Exporting Cost Rate Tables..."
915     myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
916     Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets("SourceData"))
917     oWorksheet.Name = "Cost Rate Tables"
918     oWorksheet.[A1:I1].Value = Array("PROJECT", "RESOURCE_NAME", "RESOURCE_TYPE", "ENTERPRISE", "RATE_TABLE", "EFFECTIVE_DATE", "STANDARD_RATE", "OVERTIME_RATE", "PER_USE_COST")
919     lngRow = 2
920     'make compatible with master/sub projects
921     If ActiveProject.ResourceCount > 0 Then
922       For Each oResource In ActiveProject.Resources
923         With oWorksheet
924           .Cells(lngRow, 1) = oResource.Name
925           For Each oCostRateTable In oResource.CostRateTables
926             If myResourceDemand_frm.Controls(Choose(oCostRateTable.Index, "chkA", "chkB", "chkC", "chkD", "chkE")).Value = True Then
927               For Each oPayRate In oCostRateTable.PayRates
928                 .Cells(lngRow, 1) = cptRegEx(ActiveProject.Name, "[^\\/]{1,}$")
929                 .Cells(lngRow, 2) = oResource.Name
930                 .Cells(lngRow, 3) = Choose(oResource.Type + 1, "Work", "Material", "Cost")
931                 .Cells(lngRow, 4) = oResource.Enterprise
932                 .Cells(lngRow, 5) = oCostRateTable.Name
933                 .Cells(lngRow, 6) = FormatDateTime(oPayRate.EffectiveDate, vbShortDate)
934                 .Cells(lngRow, 7) = Replace(oPayRate.StandardRate, "/h", "")
935                 .Cells(lngRow, 8) = Replace(oPayRate.OvertimeRate, "/h", "")
936                 .Cells(lngRow, 9) = oPayRate.CostPerUse
937                 lngRow = .Cells(.Rows.Count, 1).End(-4162).Row + 1
938               Next oPayRate
939             End If
940           Next oCostRateTable
941         End With
942       Next oResource
943     ElseIf ActiveProject.Subprojects.Count > 0 Then
944       For Each oSubproject In ActiveProject.Subprojects
945         For Each oResource In oSubproject.SourceProject.Resources
946           With oWorksheet
947             .Cells(lngRow, 1) = oResource.Name
948             For Each oCostRateTable In oResource.CostRateTables
949               If myResourceDemand_frm.Controls(Choose(oCostRateTable.Index, "chkA", "chkB", "chkC", "chkD", "chkE")).Value = True Then
950                 For Each oPayRate In oCostRateTable.PayRates
951                   .Cells(lngRow, 1) = cptRegEx(oSubproject.SourceProject.Name, "[^\\/]{1,}$")
952                   .Cells(lngRow, 2) = oResource.Name
953                   .Cells(lngRow, 3) = Choose(oResource.Type + 1, "Work", "Material", "Cost")
954                   .Cells(lngRow, 4) = oResource.Enterprise
955                   .Cells(lngRow, 5) = oCostRateTable.Name
956                   .Cells(lngRow, 6) = FormatDateTime(oPayRate.EffectiveDate, vbShortDate)
957                   .Cells(lngRow, 7) = Replace(oPayRate.StandardRate, "/h", "")
958                   .Cells(lngRow, 8) = Replace(oPayRate.OvertimeRate, "/h", "")
959                   .Cells(lngRow, 9) = oPayRate.CostPerUse
960                   lngRow = .Cells(.Rows.Count, 1).End(-4162).Row + 1
961                 Next oPayRate
962               End If
963             Next oCostRateTable
964           End With
965         Next oResource
966       Next oSubproject
967     End If
968
969     'make it a oListObject
970     Set oListObject = oWorksheet.ListObjects.Add(1, oWorksheet.Range(oWorksheet.[A1].End(-4161), oWorksheet.[A1].End(-4121)), , 1)
971     oListObject.Name = "CostRateTables"
972     oListObject.TableStyle = ""
973     oExcel.ActiveWindow.Zoom = 85
974     oWorksheet.[A2].Select
975     oExcel.ActiveWindow.FreezePanes = True
976     oWorksheet.Columns.AutoFit
977
978   End If
979
980   'PivotTable Worksheet active by default
981   oWorkbook.Sheets("ResourceDemand").Activate
982
983   'provide user feedback
984   Application.StatusBar = "Saving the Workbook..."
985   myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
986
987 '  'save the file
988 '  '<issue49> - file exists in location
989 '  strFileName = oShell.SpecialFolders("Desktop") & "\" & Replace(oWorkbook.Name, ".xlsx", "_" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".xlsx") '<issue49>
990 '  If Dir(strFileName) <> vbNullString Then '<issue49>
991 '    If MsgBox("A file named '" & strFileName & "' already exists in this location. Replace?", vbYesNo + vbExclamation, "Overwrite?") = vbYes Then '<issue49>
992 '      Kill strFileName '<issue49>
993 '      oWorkbook.SaveAs strFileName, 51 '<issue49>
994 '      MsgBox "Saved to your Desktop:" & vbCrLf & vbCrLf & Dir(strFileName), vbInformation + vbOKOnly, "Resource Demand Exported" '<issue49>
995 '    End If '<issue49>
996 '  Else '<issue49>
997 '    oWorkbook.SaveAs strFileName, 51  '<issue49>
998 '  End If '</issue49>
999
1000  If blnFiscal Then
1001    strMsg = "Apply an efficiency factor in cell H1 of the FiscalPeriods worksheet (e.g., 1 FTE = 85%)." & vbCrLf & vbCrLf
1002    strMsg = strMsg & "To account for calendar exceptions:" & vbCrLf
1003    strMsg = strMsg & "- use Calendar Details feature to export calendar exceptions;" & vbCrLf
1004    strMsg = strMsg & "- for recurring exceptions, be sure to select 'detailed';" & vbCrLf
1005    strMsg = strMsg & "- expand recurring exceptions to show full list of Start dates;" & vbCrLf
1006    strMsg = strMsg & "- copy list of 'Start' dates and paste into Exceptions List on FiscalPeriods sheet;" & vbCrLf
1007    strMsg = strMsg & "- activate ResourceDemand or PivotChart sheet and Refresh Pivot data" & vbCrLf & vbCrLf
1008    strMsg = strMsg & "(Take a screen shot of these instructions, if needed.)"
1009    MsgBox strMsg, vbInformation + vbOKOnly, "Next Actions:"
1010    oWorkbook.Sheets("FiscalPeriods").Activate
1011    oWorkbook.Sheets("FiscalPeriods").[E2].Select
1012  End If
1013
1014  MsgBox "Export Complete", vbInformation + vbOKOnly, "Staffing Profile"
1015
1016  Application.StatusBar = "Complete."
1017  myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
1018
1019  oExcel.Visible = True
1020  Application.ActivateMicrosoftApp pjMicrosoftExcel
1021
exit_here:
1023  On Error Resume Next
1024  If Not oExcel Is Nothing Then oExcel.Visible = True
1025  Application.StatusBar = ""
1026  myResourceDemand_frm.lblStatus.Caption = "Ready..."
1027  cptSpeed False
1028  Set oAssignment = Nothing
1029  Set oCalendar = Nothing
1030  Set oChart = Nothing
1031  Set oCostRateTable = Nothing
1032  Set oEstimates = Nothing
1033  Set oExcel = Nothing
1034  Set oException = Nothing
1035  Set oListObject = Nothing
1036  Set oPayRate = Nothing
1037  Set oPivotChartTable = Nothing
1038  Set oPivotTable = Nothing
1039  Set oRange = Nothing
1040  Set oRecordset = Nothing
1041  Set oResource = Nothing
1042  Set oSettings = Nothing
1043  Set oSubproject = Nothing
1044  Set oTask = Nothing
1045  Set oTSV = Nothing
1046  Set oTSVS_AC = Nothing
1047  Set oTSVS_AW = Nothing
1048  Set oTSVS_BCWS = Nothing
1049  Set oTSVS_COST = Nothing
1050  Set oTSVS_WORK = Nothing
1051  Set oWorkbook = Nothing
1052  Set oWorksheet = Nothing
1053
1054  If Not oWorkbook Is Nothing Then oWorkbook.Close False
1055  If Not oExcel Is Nothing Then oExcel.Quit
1056  Exit Sub
err_here:
1058  Call cptHandleErr(MODULE_NAME, "cptExportResourceDemand", Err, Erl)
1059  On Error Resume Next
1060  Resume exit_here
1061
End Sub

Sub cptShowExportResourceDemand_frm()
1065  'objects
1066  Dim myResourceDemand_frm As cptResourceDemand_frm
1067  Dim rst As ADODB.Recordset
1068  Dim rstResources As Object 'ADODB.Recordset
1069  Dim objProject As Object
1070  Dim rstFields As Object 'ADODB.Recordset
1071  'strings
1072  Dim strDir As String
1073  Dim strNonLabor As String
1074  Dim strBaseline As String
1075  Dim strCostSets As String
1076  Dim strCosts As String
1077  Dim strWeeks As String
1078  Dim strMonths As String
1079  Dim strWeekday As String
1080  Dim strMissing As String
1081  Dim strFieldName As String
1082  Dim strFileName As String
1083  'longs
1084  Dim lngFile As Long
1085  Dim lngResourceCount As Long
1086  Dim lngResource As Long
1087  Dim lngField As Long
1088  Dim lngItem As Long
1089  'integers
1090  'booleans
1091  Dim blnErrorTrapping As Boolean
1092  Dim blnFiscalCalendarExists As Boolean
1093  'variants
1094  Dim vField As Variant
1095  Dim vCostSet As Variant
1096  Dim vCostSets As Variant
1097  Dim vFieldType As Variant
1098  'dates
1099
1100  'prevent spawning
1101  If Not cptGetUserForm("cptResourceDemand_frm") Is Nothing Then Exit Sub
1102
1103  blnErrorTrapping = cptErrorTrapping
1104  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
1105  strDir = cptDir
1106
1107  'requires ms excel
1108  If Not cptCheckReference("Excel") Then
1109    MsgBox "This feature requires MS Excel.", vbCritical + vbOKOnly, "Resource Demand"
1110    GoTo exit_here
1111  End If
1112  If ActiveProject.Subprojects.Count = 0 And ActiveProject.ResourceCount = 0 Then
1113    MsgBox "This project has no resources to export.", vbExclamation + vbOKOnly, "No Resources"
1114    GoTo exit_here
1115  Else
1116    cptSpeed True
1117    lngResourceCount = ActiveProject.ResourceCount
1118    Set rstResources = CreateObject("ADODB.Recordset")
1119    rstResources.Fields.Append "RESOURCE_NAME", adVarChar, 200
1120    rstResources.Open
1121    For lngItem = 1 To ActiveProject.Subprojects.Count
1122      Set objProject = ActiveProject.Subprojects(lngItem).SourceProject
1123      Application.StatusBar = "Loading " & objProject.Name & "..."
1124      For lngResource = 1 To objProject.Resources.Count
1125        With rstResources
1126          .Filter = "[RESOURCE_NAME]='" & objProject.Resources(lngResource).Name & "'"
1127          If rstResources.RecordCount = 0 Then
1128            .AddNew Array(0), Array("'" & objProject.Resources(lngResource).Name & "'")
1129          Else
1130            Debug.Print "duplicate found"
1131          End If
1132          .Filter = ""
1133        End With
1134      Next lngResource
1135      Set objProject = Nothing
1136    Next lngItem
1137    rstResources.Close 'todo: save for later?
1138    Application.StatusBar = ""
1139    cptSpeed False
1140  End If
1141
1142  'instantiate the form
1143  Set myResourceDemand_frm = New cptResourceDemand_frm
1144  myResourceDemand_frm.lboFields.Clear
1145  myResourceDemand_frm.lboExport.Clear
1146
1147  Set rstFields = CreateObject("ADODB.Recordset")
1148  rstFields.Fields.Append "CONSTANT", adInteger
1149  rstFields.Fields.Append "CUSTOM_NAME", adVarChar, 255
1150  rstFields.Open
1151
1152  'add the 'Critical' field
1153  rstFields.AddNew Array(0, 1), Array(FieldNameToFieldConstant("Critical"), "Critical")
1154
1155  For Each vFieldType In Array("Text", "Outline Code")
1156    On Error GoTo err_here
1157    For lngItem = 1 To 30
1158      lngField = FieldNameToFieldConstant(vFieldType & lngItem) ',lngFieldType)
1159      strFieldName = CustomFieldGetName(lngField)
1160      If Len(strFieldName) > 0 Then
1161        'todo: handle duplicates if master/subprojects
1162        rstFields.AddNew Array(0, 1), Array(lngField, strFieldName)
1163        rstFields.Update
1164      End If
next_field:
1166    Next lngItem
1167  Next vFieldType
1168
1169  'get enterprise custom fields
1170  For lngField = 188776000 To 188778000
1171    If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
1172      strFieldName = Application.FieldConstantToFieldName(lngField)
1173      'todo: avoid conflicts between local and custom fields?
1174      'If rstFields.Contains(strFieldName) Then
1175      '  MsgBox "An Enterprise Field named '" & strFieldName & "' conflicts with a local custom field of the same name. The local field will be ignored.", vbExclamation + vbOKOnly, "Conflict"
1176        'rstFields.Remove Application.FieldConstantToFieldName(lngField)
1177      'End If
1178      rstFields.AddNew Array(0, 1), Array(lngField, strFieldName)
1179      rstFields.Update
1180    End If
next_field1:
1182  Next lngField
1183
1184  'add fields to listbox
1185  rstFields.Sort = "CUSTOM_NAME"
1186  rstFields.MoveFirst
1187  lngItem = 0
1188  Do While Not rstFields.EOF
1189    myResourceDemand_frm.lboFields.AddItem
1190    myResourceDemand_frm.lboFields.List(lngItem, 0) = rstFields(0)
1191    If rstFields(0) > 188776000 Then
1192      myResourceDemand_frm.lboFields.List(lngItem, 1) = rstFields(1) & " (Enterprise)"
1193    Else
1194      myResourceDemand_frm.lboFields.List(lngItem, 1) = rstFields(1) & " (" & FieldConstantToFieldName(rstFields(0)) & ")"
1195    End If
1196    rstFields.MoveNext
1197    lngItem = lngItem + 1
1198  Loop
1199
1200  'save the fields to a file for fast searching
1201  If rstFields.RecordCount > 0 Then
1202    strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
1203    If Dir(strFileName) <> vbNullString Then Kill strFileName
1204    rstFields.Save strFileName, adPersistADTG
1205  End If
1206  rstFields.Close
1207
1208  'populate options and set defaults
1209  With myResourceDemand_frm
1210    .cboWeeks.AddItem "Beginning"
1211    .cboWeeks.AddItem "Ending"
1212    'allow to trigger, it populates the form
1213    .cboWeeks.Value = "Beginning"
1214    .cboWeekday = "Monday"
1215    .chkA.Value = False
1216    .chkB.Value = False
1217    .chkC.Value = False
1218    .chkD.Value = False
1219    .chkE.Value = False
1220    .chkCosts.Value = False
1221    .chkAssociatedBaseline = False
1222    .chkFullBaseline = False
1223    .cboMonths.Clear
1224    .cboMonths.AddItem
1225    .cboMonths.List(.cboMonths.ListCount - 1, 0) = 0
1226    .cboMonths.List(.cboMonths.ListCount - 1, 1) = "Calendar (Default Excel Grouping)"
1227    .cboMonths.Value = 0
1228    blnFiscalCalendarExists = cptCalendarExists("cptFiscalCalendar")
1229    If blnFiscalCalendarExists Then
1230      .cboMonths.AddItem
1231      .cboMonths.List(.cboMonths.ListCount - 1, 0) = 1
1232      .cboMonths.List(.cboMonths.ListCount - 1, 1) = "Fiscal (cptFiscalCalendar)"
1233    Else
1234      .cboMonths.Enabled = False
1235      .cboMonths.Locked = True
1236    End If
1237  End With
1238
1239  'import saved fields if exists
1240  strFileName = strDir & "\settings\cpt-export-resource-userfields.adtg"
1241  If Dir(strFileName) <> vbNullString Then
1242    Set rst = CreateObject("ADODB.Recordset")
1243    With rst
1244      .Open strFileName, , adOpenKeyset, adLockReadOnly
1245      .MoveFirst
1246      lngItem = 0
1247      Do While Not .EOF
1248        If .Fields(0) = "settings" Then
1249          'don't use it - obsolete
1250        Else
1251          If .Fields(0) >= 188776000 Then 'check enterprise field
1252            If FieldConstantToFieldName(.Fields(0)) <> Replace(.Fields(1), cptRegEx(.Fields(1), " \([A-z0-9]*\)$"), "") Then
1253              strMissing = strMissing & "- " & .Fields(1) & vbCrLf
1254              GoTo next_saved_field
1255            End If
1256          Else 'check local field
1257            If CustomFieldGetName(.Fields(0)) <> Trim(Replace(.Fields(1), cptRegEx(.Fields(1), "\([^\(].*\)$"), "")) Then
1258              'limit this check to Custom Fields
1259              If IsNumeric(Right(FieldConstantToFieldName(.Fields(0)), 1)) Then
1260                strMissing = strMissing & "- " & .Fields(1) & vbCrLf
1261                GoTo next_saved_field
1262              End If
1263            End If
1264          End If
1265          myResourceDemand_frm.lboExport.AddItem
1266          myResourceDemand_frm.lboExport.List(lngItem, 0) = .Fields(0) 'Field Constant
1267          myResourceDemand_frm.lboExport.List(lngItem, 1) = .Fields(1) 'Custom Field Name
1268          lngItem = lngItem + 1
1269        End If
next_saved_field:
1271        .MoveNext
1272      Loop
1273      .Close
1274    End With
1275  End If
1276
1277  'import saved settings
1278  With myResourceDemand_frm
1279    If Dir(strDir & "\settings\cpt-settings.ini") <> vbNullString Then
1280      cptDeleteSetting "ResourceDemand", "chkBaseline"
1281      cptDeleteSetting "ResourceDemand", "lboExport"
1282      'month
1283      strMonths = cptGetSetting("ResourceDemand", "cboMonths")
1284      If Len(strMonths) > 0 Then
1285        If CLng(strMonths) = 1 And blnFiscalCalendarExists Then
1286          .cboMonths.Value = CLng(strMonths)
1287        Else
1288          .cboMonths.Value = 0
1289        End If
1290      End If
1291      'week
1292      strWeeks = cptGetSetting("ResourceDemand", "cboWeeks")
1293      If Len(strWeeks) > 0 Then
1294        .cboWeeks.Value = strWeeks
1295      End If
1296      'weekday
1297      strWeekday = cptGetSetting("ResourceDemand", "cboWeekday")
1298      If Len(strWeekday) > 0 Then
1299        .cboWeekday.Value = strWeekday
1300      End If
1301      'costs
1302      strCosts = cptGetSetting("ResourceDemand", "chkCosts")
1303      If Len(strCosts) > 0 Then
1304        .chkCosts = CBool(strCosts)
1305      End If
1306      If .chkCosts Then
1307        strCostSets = cptGetSetting("ResourceDemand", "CostSets")
1308        If Len(strCostSets) > 0 Then
1309          If Right(strCostSets, 1) = "," Then strCostSets = Left(strCostSets, Len(strCostSets) - 1)
1310          For Each vCostSet In Split(strCostSets, ",")
1311            .Controls("chk" & Choose(CLng(vCostSet + 1), "A", "B", "C", "D", "E")).Value = True
1312          Next vCostSet
1313        End If
1314      Else
1315        For Each vCostSet In Split("A,B,C,D,E", ",")
1316          .Controls("chk" & vCostSet).Value = False
1317          .Controls("chk" & vCostSet).Enabled = False
1318        Next vCostSet
1319      End If
1320      'baseline
1321      strBaseline = cptGetSetting("ResourceDemand", "chkAssociatedBaseline")
1322      If Len(strBaseline) > 0 Then
1323        .chkAssociatedBaseline = CBool(strBaseline)
1324      End If
1325      strBaseline = cptGetSetting("ResourceDemand", "chkFullBaseline")
1326      If Len(strBaseline) > 0 Then
1327        .chkFullBaseline = CBool(strBaseline)
1328      End If
1329      'non-labor
1330      cptDeleteSetting "ResourceDemand", "chkNonLabor" 'obsolete setting
1331    End If
1332    .Caption = "Export Resource Demand (" & cptGetVersion(MODULE_NAME) & ")"
1333    If Len(strMissing) > 0 Then
1334      If UBound(Split(strMissing, vbCrLf)) > 10 Then
1335        lngFile = FreeFile
1336        strFileName = Environ("tmp") & "\cpt-resourcedemand-missing-fields.txt"
1337        Open strFileName For Output As #lngFile
1338        Print #lngFile, "The following saved fields do not exist in this project:"
1339        Print #lngFile, strMissing
1340        Close #lngFile
1341        ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
1342        MsgBox "There are " & UBound(Split(strMissing, vbCrLf)) & " saved fields that do not exist in this project.", vbCritical + vbOKOnly, "Saved Settings"
1343      Else
1344        MsgBox "The following saved fields do not exist in this project:" & vbCrLf & strMissing, vbInformation + vbOKOnly, "Saved Settings"
1345      End If
1346    End If
1347    .Show 'False
1348  End With
1349
exit_here:
1351  On Error Resume Next
1352  Unload myResourceDemand_frm
1353  Set myResourceDemand_frm = Nothing
1354  Set rst = Nothing
1355  If rstResources.State Then rstResources.Close
1356  Set rstResources = Nothing
1357  Set objProject = Nothing
1358  If rstFields.State Then rstFields.Close
1359  Set rstFields = Nothing
1360  Exit Sub
1361
err_here:
1363  If Err.Number = 1101 Or Err.Number = 1004 Then
1364    Err.Clear
1365    Resume next_field
1366  Else
1367    Call cptHandleErr(MODULE_NAME, "cptShowExportResourceDemand_frm", Err, Erl)
1368    Resume exit_here
1369  End If
1370
End Sub

Function cptGetFiscalMonthOfDay(dtDate As Date, vFiscal As Variant)
1374  Dim lngItem As Long
1375  For lngItem = 0 To UBound(vFiscal, 2)
1376    If vFiscal(0, lngItem) >= dtDate Then
1377      cptGetFiscalMonthOfDay = vFiscal(1, lngItem)
1378      Exit Function
1379    End If
1380  Next lngItem
1381  cptGetFiscalMonthOfDay = ""
End Function
