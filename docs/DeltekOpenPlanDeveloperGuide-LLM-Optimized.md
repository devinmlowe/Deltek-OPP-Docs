# Deltek Open Plan® Developer's Guide - LLM Reference
*Optimized for AI Assistant Context*

## Quick Start

**Create Application:**
```vb
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")
OPApp.Login  'Required for 8.x+
```

**Open Project:**
```vb
Set proj = OPApp.FileOpen("path\project.opp")
proj.TimeAnalyze
proj.Save
```

**Work with Activities:**
```vb
Set acts = proj.Activities
Set act = acts.Item("ACT001")  'By ID
act.PercentComplete = 50
act.SetField "USER1", "value"
```

## Data Types
- **Str**: String
- **Int**: Integer/Long
- **Bool**: Boolean
- **Dbl**: Double
- **Date**: Date
- **R/W**: Read/Write
- **RO**: Read Only

---

## OBJECTS & COLLECTIONS

### OPCreateApplication33 (Application Root)
**Props:** EnableUndo(Bool,R/W), Height(Int,R/W), SilentMode(Bool,R/W), Width(Int,R/W), XPosition(Int,R/W), YPosition(Int,R/W)

**Methods:** ActiveProject, ActiveView, ApplyFilter, Calendars, Codes, Conceal, CreateBrowserView, Disable, Enable, FileCabinet, FileNew, FileOpen, FileOpenEx, FilePrint, FilePrintPreview, FilePrintSetup, FileRename, GeneralExport, GeneralImport, GetLastSecurityValidation, Login, Maximize, Minimize, PageSetup, Projects, ReportingCalendars, Resources, Restore, SetPrinter, SetResourceSelection, Show, SysDir, User, Version, WebWindows, WindowMinimizeAll, WindowTile, WindowTileVertical, WorkDir

### OPProject
**Props:** AccessMode(Str,RO), AutoAnalyze(Bool,R/W), AutoProgActBasedOn(Str,R/W), AutoProgActComplete(Bool,R/W), AutoProgActFilter(Str,R/W), AutoProgActInProgress(Bool,R/W), AutoProgActivity(Bool,R/W), AutoProgActProgressType(Str,R/W), AutoProgActSetPPC(Bool,R/W), AutoProgResEndDate(Date,R/W), AutoProgResource(Bool,R/W), AutoProgResStartDate(Date,R/W), CalcCostActual(Bool,R/W), CalcCostBasedOn(Str,R/W), CalcCostBudget(Bool,R/W), CalcCostEarnedValues(Bool,R/W), CalcCostEscalated(Bool,R/W), CalcCostIncludeChildValues(Bool,R/W), CalcCostRemaining(Bool,R/W), Calendar(Str,R/W), Client(Str,R/W), Company(Str,R/W), CurrentActivity(Str,RO), DateFormat(Str,R/W), DefaultActivityCalendar(Str,R/W), DefaultActivityType(Str,R/W), DefaultAuxiliaryAccessMode(Str,R/W), DefaultDurationChar(Str,RO), DefaultDurationUnit(Str,R/W), DefaultProjectAccessMode(Str,R/W), DefaultRelationshipCalendar(Str,R/W), Description(Str,R/W), EarlyFinishDate(Date,R/W), FileName(Str,RO), FiscalCalendar(Str,R/W), HardZeroes(Bool,R/W), Height(Int,R/W), InProgressPriority(Str,R/W), LateFinishDate(Date,R/W), Manager(Str,R/W), MinimumCalcDurationUnit(Str,R/W), MinutesPerDay(Int,R/W), MinutesPerDefaultUnit(Int,RO), MinutesPerMinDurUnit(Int,RO), MinutesPerMonth(Int,R/W), MinutesPerWeek(Int,R/W), MultipleEnd(Int,R/W), Name(Str,RO), OutOfSeqOpt(Int,R/W), Priority1Name(Str,R/W), Priority2Name(Str,R/W), Priority3Name(Str,R/W), Resource(Str,R/W), ScheduleFinishDate(Date,R/W), ScheduleMethod(Str,R/W), ScheduleTimeUnit(Str,R/W), ScheduleTimeUnitMinutes(Int,RO), Smoothing(Bool,R/W), StartDate(Date,R/W), StartView(Str,R/W), StatusDate(Date,R/W), TargetCost(Dbl,R/W), TargetFinishDate(Date,R/W), TargetFinishType(Str,R/W), TargetStartDate(Date,R/W), Width(Int,R/W), XPosition(Int,R/W), YPosition(Int,R/W)

**Methods:** Activities, AutoProgress, AutoProgressEx, Barcharts, BaselinesList, BatchGlobalEdits, CalculateCost, CalculateCostEx, CalculatedFields, Categories, CloseView, CodeFiles, Conceal, Costs, CreateBackup, CreateBaseline(name), DeleteBaseline(name), Disable, Enable, Filters, GeneralExport, GeneralImport, GenerateCrosstabDates, GetCrosstabDates, GetCrosstabDatesInXML, GetField(fieldname), GetFields(array), GetSelectedActivitiesArray, GetSharedModeLock, GlobalEdits, Graphs, Maximize, Minimize, Networks, Notes, Resources, ResourceSchedule(method,autoanalyze), Restore, RiskAnalyze, Rollups, Save, SaveAs(path), SetCrosstabDates, SetCrosstabDatesFromXML, SetEarnedValueCrosstabOptions, SetField(fieldname,value), SetResourceCrosstabOptions, SetSharedmodeLock, Shut, ShutWOSave, Sorts, Spreadsheets, TimeAnalyze, UpdateBaseline(name), UpdateBaselineEx, Views

### OPActivity
**Props:** ActivityLogicFlag(Str,RO), ActualFinishDate(Date,R/W), ActualStartDate(Date,R/W), BaselineFinish(Date,R/W), BaselineStart(Date,R/W), BoxHeight(Int,R/W), BoxWidth(Int,R/W), BudgetCost(Dbl,R/W), Calendar(Str,R/W), ComputedRemainingDuration(Str,RO), ComputedStatus(Str,RO), CostToDate(Dbl,R/W), Critical(Int,RO), CriticalIndex(Int,RO), DelayingResource(Str,RO), Description(Str,R/W), Duration(Str,R/W), DurationDistShape(Str,R/W), EarliestFeasible(Date,RO), EarlyFinish(Date,RO), EarlyFinishStdDev(Str,RO), EarlyStart(Date,RO), EarlyStartStdDev(Str,RO), ExpectedFinish(Date,R/W), FinishFreeFloat(Str,RO), FinishTotalFloat(Str,RO), FirstUsage(Date,RO), FreeFloat(Str,RO), FreeFloatStdDev(Str,RO), ID(Str,R/W), KeyActivityStatus(Str,R/W), LateFinish(Date,RO), LateFinishStdDev(Str,RO), LateStart(Date,RO), LateStartStdDev(Str,RO), MaxDuration(Str,R/W), MaxNoSplits(Int,R/W), MinSplitLength(Str,R/W), OptimisticDuration(Str,R/W), PercentComplete(Dbl,R/W), PessimisticDuration(Str,R/W), ProgressFlag(Str,R/W), ProgressValue(Str,R/W), RefreshData(Bool,R/W), ResourceScheduleType(Str,R/W), ScheduleActions(Str,R/W), ScheduledFinish(Date,R/W), ScheduledStart(Date,R/W), ScheduleDuration(Str,R/W), ScheduleFloat(Str,RO), SubprojectFilename(Str,R/W), SuppressRequirements(Bool,R/W), TargetFinish(Date,R/W), TargetFinishType(Str,R/W), TargetStart(Date,R/W), TargetStartType(Str,R/W), TotalFloat(Str,RO), TotalFloatStdDev(Str,RO), TotalResourceCost(Dbl,RO), Type(Str,R/W)

**Methods:** ActivityResources, Assignments, GetCurrentFields, GetField(name), GetFields(array), IsExternalOpen, Notes, Predecessors, Remove, ResourceUsages, Selected, SetCurrentFields(array), SetField(name,value), Steps

### OPActivities (Collection)
**Props:** Count(Int,RO), Filter(Str,R/W), Sort(Str,R/W)
**Methods:** Add(id,desc), AssignCurrentFieldSet, Categories, Fields, GetCalculatedFieldString, GetFilterString, Item(index/id), Remove(index/id), SetAssignmentFields, SetCalculatedFieldTo, SetCollectionGrowth, SetCostFields, SetFilterTo, SetPredFields, SetRiskFields, SetSortFields, SetSortTo, SetStepFields, SetSuccFields, SetUsageFields, Steps

### OPActivityResource
**Props:** ActivityID(Str,RO), Class(Str,R/W), Description(Str,R/W), ID(Str,RO), RefreshData(Bool,R/W), RollUp(Str,R/W), Suppress(Bool,R/W), Threshold(Str,R/W), Type(Str,R/W), UnitCost(Dbl,R/W), Units(Dbl,R/W)
**Methods:** Availabilities, GetCurrentFields, GetEarnedValueCrosstabData, GetEarnedValueCrosstabDataInXML, GetField(name), GetFields(array), GetResourceCrosstabData, GetResourceCrosstabDataInXML, GetResourceDateArray, Notes, Remove, SetCurrentFields, SetField(name,value)

### OPAssignment
**Props:** ActivityID(Str,R/W), AlternateResourceID(Str,R/W), Level(Dbl,R/W), LevelType(Str,R/W), RefreshData(Bool,R/W), Remaining(Dbl,R/W), ResourceID(Str,R/W), ResourceOffset(Str,R/W), ResourcePeriod(Str,R/W)
**Methods:** GetCurrentFields, GetEarnedValueCrosstabData, GetEarnedValueCrosstabDataInXML, GetField(name), GetFields(array), GetResourceCrosstabData, GetResourceCrosstabDataInXML, GetResourceDateArray, Remove, SetCurrentFields, SetField(name,value)

### OPAssignments (Collection)
**Props:** Count(Int,RO)
**Methods:** Add, AssignCurrentFieldSet, Fields, Item(index), Remove(index), SetCollectionGrowth

### OPAvailability
**Props:** Calendar(Str,R/W), Level(Dbl,R/W), RefreshData(Bool,R/W), Resource(Str,RO), StartDate(Date,R/W), StopDate(Date,R/W)
**Methods:** GetCurrentFields, GetField(name), GetFields(array), Remove, SetCurrentFields, SetField(name,value)

### OPAvailabilities (Collection)
**Props:** Count(Int,RO)
**Methods:** Add, AssignCurrentFieldSet, Fields, Item(index), Remove(index)

### OPPredecessor
**Props:** Calendar(Str,R/W), FreeFloat(Str,RO), ID(Str,R/W), Lag(Str,R/W), RefreshData(Bool,R/W), RelationshipType(Str,R/W), SuccessorID(Str,RO), TotalFloat(Str,RO), Turns(Int,R/W)
**Methods:** GetCurrentFields, GetField(name), GetFields(array), Remove, SetCurrentFields, SetField(name,value)

### OPPredecessors (Collection)
**Props:** Count(Int,RO), Sort(Str,R/W)
**Methods:** Add(predID), AssignCurrentFieldSet, Fields, Item(index), Remove(index), SetCollectionGrowth, SetSortFields, SetSortTo

### OPProjectResource
**Props:** Class(Str,R/W), Description(Str,R/W), ID(Str,RO), RefreshData(Bool,R/W), RollUp(Str,R/W), Suppress(Bool,R/W), Threshold(Str,R/W), Type(Str,R/W), UnitCost(Dbl,R/W), Units(Dbl,R/W)
**Methods:** Availabilities, GetCurrentFields, GetEarnedValueCrosstabData, GetEarnedValueCrosstabDataInXML, GetField(name), GetFields(array), GetResourceCrosstabData, GetResourceCrosstabDataInXML, GetResourceDateArray, Notes, Remove, SetCurrentFields, SetField(name,value)

### OPProjectResources (Collection)
**Props:** Count(Int,RO), FileName(Str,RO)
**Methods:** Add, AssignCurrentFieldSet, CalculatedFields, Categories, Fields, Filters, GlobalEdits, Item(index/id), Rollups, Sorts

### OPResourceRecord
**Props:** Class(Str,R/W), Description(Str,R/W), ID(Str,R/W), RefreshData(Bool,R/W), RollUp(Str,R/W), Suppress(Bool,R/W), Threshold(Str,R/W), Type(Str,R/W), UnitCost(Dbl,R/W), Units(Dbl,R/W)
**Methods:** AssignSkill, Availabilities, GetAvailabilityCrosstabData, GetAvailabilityCrosstabDataInXML, GetCurrentFields, GetField(name), GetFields(array), GetSkills, Notes, Remove, RemoveSkill, SetCurrentFields, SetField(name,value)

### OPResource (Collection - Resource File)
**Props:** Calendar(Str,R/W), Count(Int,RO), FileName(Str,RO), Filter(Str,R/W), Height(Int,R/W), Sort(Str,R/W), Width(Int,R/W), XPosition(Int,R/W), YPosition(Int,R/W)
**Methods:** Add, AssignCurrentFieldSet, CalculatedFields, Categories, Conceal, Disable, Enable, Fields, Filters, GenerateCrosstabDates, GetCrosstabDates, GetCrosstabDatesInXML, GetCrosstabMinutesPerDefaultUnit, GetField(name), GetFields(array), GlobalEdits, Item(index/id), Maximize, Minimize, Remove(index/id), Restore, Rollups, Save, SaveAs(path), SetCollectionGrowth, SetCrosstabDates, SetCrosstabDatesFromXML, SetCrosstabMinutesPerDefaultUnit, SetField(name,value), SetResourceCrosstabOptions, Shut, ShutWOSave, Sorts

### OPResources (Collection)
**Props:** Count(Int,RO)
**Methods:** Fields, Item(index/name)

### OPResourceUsage
**Props:** ActivityID(Str,R/W), AlternateResourceID(Str,R/W), Cost(Dbl,R/W), EscalatedCost(Dbl,R/W), FinishDate(Date,R/W), RefreshData(Bool,R/W), ResourceID(Str,R/W), StartDate(Date,R/W), Used(Dbl,R/W)
**Methods:** GetCurrentFields, GetField(name), GetFields(array), Remove, SetCurrentField, SetField(name,value)

### OPResourceUsages (Collection)
**Props:** Count(Int,RO)
**Methods:** Add, Item(index)

### OPCodeRecord
**Props:** Code(Str,R/W), Description(Str,R/W), RefreshData(Bool,R/W)
**Methods:** GetCurrentFields, GetField(name), GetFields(array), Notes, Remove, SetCurrentFields, SetField(name,value)

### OPCode (Collection - Code File)
**Props:** Count(Int,RO), Filename(Str,RO), Height(Int,R/W), Width(Int,R/W), XPosition(Int,R/W), YPosition(Int,R/W)
**Methods:** Add, AssignCurrentFieldSet, Categories, Conceal, Disable, Enable, Fields, GetField(name), GetFields(array), Item(index/code), Maximize, Minimize, Remove(index/code), Restore, Rollups, Save, SaveAs(path), SetCollectionGrowth, SetField(name,value), Shut, ShutWOSave

### OPCodes (Collection)
**Props:** Count(Int,RO)
**Methods:** Fields, Item(index/name)

### OPProjectCode (Collection)
**Props:** Count(Int,RO), FieldName(Str,RO), FileName(Str,RO), Height(Int,R/W), Width(Int,R/W), XPosition(Int,R/W), YPosition(Int,R/W)
**Methods:** Add, Categories, Conceal, Disable, Enable, Item(index/code), Maximize, Minimize, Remove(index/code), Restore, Rollups, Save, SaveAs(path), SetCollectionGrowth, Shut, ShutWOSave

### OPProjectCodes (Collection)
**Props:** Count(Int,RO)
**Methods:** Add, Item(index/name), Remove(index/name)

### OPCalendarRecord
**Methods:** Date(date), ExtraWorkDays, GetField(name), GetFields(array), GetStandardDays, Holidays, Name, SetField(name,value), SetStandardDays, StandardDay(daynum)

### OPCalendar (Collection - Calendar File)
**Props:** Count(Int,RO), Filename(Str,RO), Height(Int,R/W), Width(Int,R/W), XPosition(Int,R/W), YPosition(Int,R/W)
**Methods:** Add, Conceal, Copy, Disable, Enable, Fields, GetField(name), GetFields(array), Item(index/name), Maximize, Minimize, Remove(index/name), Restore, Save, SaveAs(path), SetCollectionGrowth, SetField(name,value), Shut, ShutWOSave

### OPCalendars (Collection)
**Props:** Count(Int,RO)
**Methods:** Fields, Item(index/name)

### OPDate
**Props:** Date(Date,RO), Work(Bool,R/W)
**Methods:** Shifts

### OPExtraWorkDays (Collection)
**Props:** Count(Int,RO)
**Methods:** Add(date), Item(index), Remove(index)

### OPHolidays (Collection)
**Props:** Count(Int,RO)
**Methods:** Add(date), AddAll(array), GetAll, Item(index), Remove(index)

### OPStandardDay
**Props:** Work(Bool,R/W)
**Methods:** Shifts

### OPShift
**Props:** StartTime(Str,R/W), StopTime(Str,R/W)

### OPShifts (Collection)
**Props:** Count(Int,RO)
**Methods:** Add(start,stop), AddAll(array), GetAll, Item(index), Remove(index), RemoveAll

### OPCost
**Props:** ActivityID(Str,R/W), ActualCost(Dbl,R/W), ActualQty(Dbl,R/W), EndDate(Date,R/W), RefreshData(Bool,R/W), ResourceID(Str,R/W), StartDate(Date,R/W)
**Methods:** GetCurrentFields, GetField(name), GetFields(array), Remove, SetCurrentFields, SetField(name,value)

### OPCosts (Collection)
**Props:** Count(Int,RO)
**Methods:** Add, AssignCurrentFieldSet, Item(index), Remove(index), SetCollectionGrowth

### OPNote
**Props:** Category(Str,RO), ModifiedBy(Str,RO), ModifiedDate(Date,RO), NoteText(Str,R/W)

### OPNotes (Collection)
**Props:** Count(Int,RO)
**Methods:** Add(category), Item(index), Remove(index)

### OPBaselines (Collection)
**Props:** Count(Int,RO), Description(Str,RO), Filter(Str,RO), Name(Str,RO), Selected(Bool,RO), Sort(Str,RO), Type(Str,RO)
**Methods:** Item(index/id)

### OPBaselinesList (Collection)
**Props:** Count(Int,RO)
**Methods:** Item(index/name), Select(name)

### OPFilter
**Props:** Expression(Str,R/W), Name(Str,RO), TableName(Str,RO)

### OPFilters (Collection)
**Props:** Count(Int,RO)
**Methods:** Add(name,expr,table), Item(index/name), Remove(index/name)

### OPSort
**Props:** Expression(Str,R/W), Name(Str,RO), TableName(Str,RO)

### OPSorts (Collection)
**Props:** Count(Int,RO)
**Methods:** Add(name,expr,table), Item(index/name), Remove(index/name)

### OPCalculatedField
**Props:** CalcAcross(Bool,R/W), Decimals(Int,R/W), Expression(Str,R/W), Name(Str,RO), ResultType(Str,RO), TableName(Str,RO)

### OPCalculatedFields (Collection)
**Props:** Count(Int,RO)
**Methods:** Add(name,expr,table), Item(index/name), Remove(index/name)

### OPGlobalEdit
**Props:** ApplyToField(Str,R/W), Expression(Str,R/W), Filter(Str,R/W), Name(Str,RO), TableName(Str,RO)
**Methods:** Apply

### OPGlobalEdits (Collection)
**Props:** Count(Int,RO)
**Methods:** Add(name,expr,table), Item(index/name), Remove(index/name)

### OPBatchGlobalEdit
**Props:** Name(Str,RO)
**Methods:** Apply

### OPBatchGlobalEdits (Collection)
**Props:** Count(Int,RO)
**Methods:** Item(index/name)

### OPRollup
**Props:** Name(Str,RO)
**Methods:** Apply, Remove

### OPRollups (Collection)
**Props:** Count(Int,RO)
**Methods:** Add, Item(index/name)

### OPCategory
**Props:** Name(Str,RO)

### OPCategories (Collection)
**Props:** Count(Int,RO)
**Methods:** Add(name), Item(index/name), Remove(index/name)

### OPSkill
**Methods:** Description, Name, UnitCost

### OPSkills (Collection)
**Props:** Count(Int,RO)
**Methods:** Add, Item(index), Remove(index)

### OPStep
**Props:** Complete(Bool,R/W), PercentComplete(Dbl,R/W)

### OPView
**Methods:** Activate, CollapseAll, Conceal, Description, Disable, Enable, ExpandAll, FileName, Maximize, Minimize, Name, RefreshFilterandSort, Restore, SetDateAutoScale, SetDateScaleOption, SetNumberOfAxes, SetPrintDateOption, SetPrintDateRange, SetPrintExpandColumns, SetPrintHeadingColorsOff, SetPrintLineNumbers, SetPrintMargins, SetPrintNotes, SetViewHorizontalPage, SetViewLegend, SetViewResourceData, SetViewSpreadsheet, Snapshot, ViewCategory, ViewClass

### OPViews (Collection)
**Props:** Count(Int,RO)
**Methods:** ActivateByFilename(name), ActivateByFileNameEx(name,mode), AddView(name), Item(index/name)

### OPBarcharts (Collection)
**Props:** Count(Int,RO)
**Methods:** ActivateByFilename(name), AddView(name), Item(index/name)

### OPNetworks (Collection)
**Props:** Count(Int,RO)
**Methods:** ActivateByFilename(name), AddView(name), Item(index/name)

### OPSpreadsheets (Collection)
**Props:** Count(Int,RO)
**Methods:** ActivateByFieldname(name), AddView(name), Item(index/name)

### OPGraphs (Collection)
**Props:** Count(Int,RO)
**Methods:** ActivateByFilename(name), AddView(name), Item(index/name)

### OPFileCabinet
**Props:** Height(Int,R/W), Width(Int,R/W), XPosition(Int,R/W), YPosition(Int,R/W)
**Methods:** Barcharts, Calendars, Codes, Conceal, Disable, Enable, Graphs, Maximize, Minimize, Networks, Projects, Resources, Restore, Spreadsheets, Views

### OPIcon
**Methods:** Activate, FileName, GetField(name), GetFields(array), Name, SetField(name,value)

### OPFCView
**Methods:** Description, Name

### OPFCProjects, OPFCResources, OPFCCalendars, OPFCCodes (Collections)
**Props:** Count(Int,RO)
**Methods:** Fields, Item(index)

### OPFCBarcharts, OPFCNetworks, OPFCSpreadsheets, OPFCGraphs, OPFCViews (Collections)
**Props:** Count(Int,RO)
**Methods:** Item(index)

### OPWebWindow
**Props:** QueryString(Str,RO), Title(Str,R/W), URLAddress(Str,R/W)
**Methods:** Close, Conceal, Disable, Enable, GetPosition, Maximize, Minimize, Refresh, Restore, SetFocus, SetPosition

### OPWebWindows (Collection)
**Props:** Count(Int,RO)
**Methods:** Add, Item(index)

### OPReportingCalendarRecord
**Props:** Date(Date,RO), Label(Str,R/W)

### OPReportingCalendar (Collection)
**Props:** Count(Int,RO), FileName(Str,RO)
**Methods:** Add, Item(index), Remove(index)

### OPReportingCalendars (Collection)
**Props:** Count(Int,RO)
**Methods:** Item(index/name)

### OPField
**Props:** DBType(Str,RO), FieldName(Str,RO), FieldType(Str,RO), IsEditable(Bool,RO), Length(Int,RO), Scale(Int,RO), TableName(Str,RO), Type(Str,RO), UserName(Str,RO)

### OPFields (Collection)
**Props:** Count(Int,RO)
**Methods:** Item(index/name)

### OPAccessControlObject
**Methods:** GetRights, SetRights

### OPAccessControlList (Collection)
**Props:** Count(Int,RO)
**Methods:** Add, Item(index), Remove(index)

### OPProjects (Collection)
**Props:** Count(Int,RO)
**Methods:** Fields, Item(index/name)

---

## KEY ENUMERATIONS

**RelationshipType:** "FS" (Finish to Start), "SS" (Start to Start), "FF" (Finish to Finish), "SF" (Start to Finish)

**ProgressFlag:** "Planned", "Remaining Duration", "Percent Complete", "Elapsed Duration", "Complete"

**ResourceScheduleType:** "Time Limited", "Resource Limited"

**CalcCostBasedOn:** "EARLY", "LATE", "SCHEDULE", "BASELINE"

**AutoProgActProgressType:** "REMAINING", "ELAPSED", "PERCENT", "EXPECTED"

**DurationDistShape:** "Beta", "None", "Normal", "Triangular", "Uniform"

**Duration Format:** Number + Unit (M=months, W=weeks, D=days, H=hours, T=minutes). Ex: "10D", "2W"

**Resource Class:** "Labor", "Material", "Other Direct Costs", "Subcontract"

---

## COMMON PATTERNS

**Iterate Collection:**
```vb
For i = 1 To collection.Count
    Set obj = collection.Item(i)
    'Work with obj
Next i
```

**Apply Filter:**
```vb
acts.Filter = "EARLY_START < #1/1/2024#"
```

**Set Sort:**
```vb
acts.Sort = "EARLY_START"
```

**Add Predecessor:**
```vb
Set pred = act.Predecessors.Add("ACT001")
pred.RelationshipType = "FS"
pred.Lag = "2D"
```

**GetField/SetField (User-defined fields):**
```vb
value = act.GetField("USER1")
act.SetField "USER1", "New Value"
```

**GetFields (Multiple fields):**
```vb
Dim fields(2) As String
fields(0) = "ACT_ID"
fields(1) = "DESCRIPTN"
fields(2) = "ORIG_DUR"
values = act.GetFields(fields)
```

**Add Resource Assignment:**
```vb
Set assign = act.Assignments.Add
assign.ResourceID = "RES001"
assign.Level = 1.0
assign.ResourcePeriod = "5D"
```

**Crosstab Data:**
```vb
proj.GenerateCrosstabDates startDate, endDate, "W"
data = resource.GetResourceCrosstabData()
xmlData = resource.GetResourceCrosstabDataInXML()
```

**Time Analysis:**
```vb
proj.TimeAnalyze
```

**Resource Schedule:**
```vb
proj.ResourceSchedule "Time Limited", True  'autoanalyze=True
proj.ResourceSchedule "Resource Limited", False
```

**Cost Calc:**
```vb
proj.CalcCostBudget = True
proj.CalcCostBasedOn = "EARLY"
proj.CalculateCost
```

**Auto Progress:**
```vb
proj.AutoProgActivity = True
proj.AutoProgActBasedOn = "EARLY"
proj.AutoProgActInProgress = True
proj.AutoProgress
```

**Create Baseline:**
```vb
proj.CreateBaseline "Baseline1"
proj.UpdateBaseline "Baseline1"
```

**Work with Baseline:**
```vb
Set baselines = proj.BaselinesList
Set baseline = baselines.Item("Baseline1")
baselines.Select "Baseline1"
Set baselineActs = baseline
For i = 1 To baselineActs.Count
    Set baselineAct = baselineActs.Item(i)
    'Compare with current
Next i
```

**Shared Mode:**
```vb
proj.SetSharedmodeLock True
act.RefreshData = True  'Refresh from DB
value = act.ID  'Gets fresh data
proj.SetSharedmodeLock False
```

**Window Management:**
```vb
OPApp.Show  'Show application
proj.Maximize  'Maximize project
view.Activate  'Activate view
view.Conceal  'Hide view
```

**Import/Export:**
```vb
proj.GeneralExport exportParams
proj.GeneralImport importParams
OPApp.GeneralExport exportParams
OPApp.GeneralImport importParams
```

---

## SECURITY (8.x+)

**Required:** Call `OPApp.Login` immediately after CreateObject

**Check Validation:**
```vb
result = OPApp.GetLastSecurityValidation
```

**User Info:**
```vb
username = OPApp.User
```

---

## UNSUPPORTED (8.x+)

**Objects:** OPFCProcesses, OPFCStructures
**Properties:** OPPFormat, ResidueFile
**Methods:** FileOpenSpecialOpenPlan, OpenSpecialODBC, GlobalDir, Rights, Processes, Structures

---

## SYSTEM DIRECTORIES

**System Dir:** `path = OPApp.SysDir`
**Work Dir:** `path = OPApp.WorkDir`
**Version:** `ver = OPApp.Version`

---

## ERROR HANDLING

```vb
On Error Resume Next
'operation
If Err.Number <> 0 Then
    'Handle error
    Err.Clear
End If
On Error GoTo 0
```

---

## DATE FORMATS

Use Project.DateFormat property. Common formats:
- "MM/DD/YY"
- "DD-MMM-YYYY"
- "YYYY-MM-DD"

Refer to Open Plan documentation for full format codes.

---

*This optimized reference contains all API information from the full guide in <20% of the tokens. For detailed descriptions and examples, consult the full DeltekOpenPlanDeveloperGuide.md file.*
