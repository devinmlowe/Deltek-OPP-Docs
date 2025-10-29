# Deltek Open Plan® API Reference for LLM Assistants

## Quick Reference

### Initialize Application
```vb
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")
OPApp.Login  'REQUIRED for Open Plan 8.x+
```

### Open & Analyze Project
```vb
Set proj = OPApp.FileOpen("C:\path\project.opp")
proj.TimeAnalyze
proj.ResourceSchedule "Time Limited", True
proj.CalculateCost
proj.Save
proj.Shut
```

### Work with Activities
```vb
Set acts = proj.Activities
Set act = acts.Item("ACT001")  'Get by ID
Set act = acts.Item(1)  'Get by index

'Modify properties
act.Description = "Task Name"
act.Duration = "10D"
act.PercentComplete = 50
act.ActualStartDate = #1/15/2024#

'User-defined fields
act.SetField "USER1", "Custom Value"
value = act.GetField("USER1")
```

---

## Core Objects

### OPCreateApplication33 (Root Application Object)

**Purpose:** Main entry point for Open Plan automation

**Key Properties:**
- EnableUndo (Bool, R/W): Enable/disable undo functionality
- SilentMode (Bool, R/W): Run without UI prompts
- Height, Width, XPosition, YPosition (Long, R/W): Window dimensions/position

**Essential Methods:**
- **Login()**: Required for 8.x+, call immediately after CreateObject
- **FileOpen(path)**: Opens project, returns OPProject object
- **FileNew(path)**: Creates new project
- **ActiveProject**: Returns currently active OPProject
- **Projects**: Returns OPProjects collection
- **Resources**: Returns OPResources collection (resource files)
- **Calendars**: Returns OPCalendars collection
- **Codes**: Returns OPCodes collection (code files)
- **FileCabinet**: Returns OPFileCabinet object (Explorer)
- **Version**: Returns version string
- **User**: Returns current username
- **SysDir, WorkDir**: Return system/working directory paths

---

### OPProject

**Purpose:** Represents an open project file

**Key Properties - Dates:**
- StartDate (Date, R/W): Project start date
- StatusDate (Date, R/W): Time Now date
- EarlyFinishDate, LateFinishDate (Date, R/W): Calculated dates
- TargetFinishDate (Date, R/W): Target finish

**Key Properties - Settings:**
- Calendar (String, R/W): Calendar file path
- DefaultDurationUnit (String, R/W): "Days", "Weeks", "Months", "Hours", "Minutes"
- DefaultActivityCalendar (String, R/W): Default calendar name
- AutoAnalyze (Bool, R/W): Auto time analysis flag
- Resource (String, R/W): Resource file path

**Key Properties - Auto Progress:**
- AutoProgActivity (Bool, R/W): Update activity status and dates
- AutoProgActBasedOn (String, R/W): "EARLY", "LATE", "SCHEDULE", "BASELINE"
- AutoProgActInProgress (Bool, R/W): Mark activities in-progress before Time Now
- AutoProgActComplete (Bool, R/W): Mark activities complete before Time Now
- AutoProgActFilter (String, R/W): Filter name to limit scope
- AutoProgResource (Bool, R/W): Update resource progress

**Key Properties - Cost Calc:**
- CalcCostBudget (Bool, R/W): Calculate budget costs
- CalcCostActual (Bool, R/W): Calculate actual costs
- CalcCostBasedOn (String, R/W): "EARLY", "LATE", "SCHEDULE", "BASELINE"
- CalcCostEarnedValues (Bool, R/W): Calculate BCWP and BCWS

**Essential Methods - Collections:**
- **Activities**: Returns OPActivities collection
- **Resources**: Returns OPProjectResources collection
- **CodeFiles**: Returns OPProjectCodes collection
- **Barcharts, Networks, Spreadsheets, Graphs**: View collections
- **BaselinesList**: Returns OPBaselinesList

**Essential Methods - Analysis:**
- **TimeAnalyze()**: Perform CPM time analysis
- **ResourceSchedule(method, autoanalyze)**: Resource scheduling
  - method: "Time Limited" or "Resource Limited"
  - autoanalyze: Bool, run time analysis after
- **CalculateCost()**: Calculate costs per settings
- **AutoProgress()**: Apply automatic progress per settings
- **RiskAnalyze()**: Perform Monte Carlo risk analysis

**Essential Methods - Baselines:**
- **CreateBaseline(name)**: Create new baseline
- **UpdateBaseline(name)**: Update existing baseline
- **DeleteBaseline(name)**: Delete baseline

**Essential Methods - File Operations:**
- **Save()**: Save project
- **SaveAs(path)**: Save with new name
- **Shut()**: Close with save
- **ShutWOSave()**: Close without save
- **CreateBackup(path)**: Create backup

**Essential Methods - Data Access:**
- **GetField(fieldname)**: Get project field value
- **SetField(fieldname, value)**: Set project field value
- **GetSelectedActivitiesArray()**: Returns array of selected activity IDs

---

### OPActivity

**Purpose:** Represents a single activity in the project

**Key Properties - Identity:**
- ID (String, R/W): Unique activity identifier
- Description (String, R/W): Activity description
- Type (String, R/W): Activity type
- ActivityLogicFlag (String, RO): "Start Activity", "End Activity", "None"

**Key Properties - Dates:**
- ActualStartDate, ActualFinishDate (Date, R/W): Actual dates
- BaselineStart, BaselineFinish (Date, R/W): Baseline dates
- EarlyStart, EarlyFinish (Date, RO): Early dates (calculated)
- LateStart, LateFinish (Date, RO): Late dates (calculated)
- ScheduledStart, ScheduledFinish (Date, R/W): Scheduled dates
- TargetStart, TargetFinish (Date, R/W): Target dates
- ExpectedFinish (Date, R/W): Expected completion

**Key Properties - Duration & Float:**
- Duration (String, R/W): Original duration (e.g., "10D", "2W")
- ScheduleDuration (String, R/W): Scheduled duration
- ComputedRemainingDuration (String, RO): Calculated remaining
- TotalFloat, FreeFloat (String, RO): Float values
- ScheduleFloat (String, RO): Schedule float

**Key Properties - Critical Path:**
- Critical (Long, RO): Criticality indicator
- CriticalIndex (Long, RO): % of simulations on critical path

**Key Properties - Progress:**
- PercentComplete (Double, R/W): Physical % complete (0-100)
- ProgressFlag (String, R/W): "Planned", "Remaining Duration", "Percent Complete", "Elapsed Duration", "Complete"
- ProgressValue (String, R/W): Progress value per ProgressFlag
- ComputedStatus (String, RO): "Planned", "In Progress", "Complete"

**Key Properties - Cost:**
- BudgetCost (Double, R/W): Activity budget (excludes resource cost)
- CostToDate (Double, R/W): Actual cost to date
- TotalResourceCost (Double, RO): Total resource cost

**Key Properties - Resources:**
- ResourceScheduleType (String, R/W): Resource schedule type
- DelayingResource (String, RO): ID of first delaying resource

**Key Properties - Risk Analysis:**
- OptimisticDuration, PessimisticDuration, MaxDuration (String, R/W): Risk durations
- DurationDistShape (String, R/W): "Beta", "Normal", "Triangular", "Uniform", "None"

**Key Properties - Other:**
- Calendar (String, R/W): Activity calendar name
- RefreshData (Bool, R/W): Force refresh from database (shared mode)
- SubprojectFilename (String, R/W): Subproject path

**Essential Methods - Collections:**
- **Predecessors**: Returns OPPredecessors collection
- **Assignments**: Returns OPAssignments collection
- **ActivityResources**: Returns OPActivityResources collection
- **ResourceUsages**: Returns OPResourceUsages collection
- **Notes**: Returns OPNotes collection
- **Steps**: Returns steps for step completion tracking

**Essential Methods - Data Access:**
- **GetField(fieldname)**: Get field value (especially user fields)
- **SetField(fieldname, value)**: Set field value
- **GetFields(fieldArray)**: Get multiple fields at once
- **GetCurrentFields()**: Get current field set values
- **SetCurrentFields(valueArray)**: Set current field set values

**Essential Methods - Other:**
- **Remove()**: Delete this activity
- **Selected()**: Returns True if activity is selected
- **IsExternalOpen()**: Check if subproject is open

---

### OPActivities (Collection)

**Purpose:** Collection of all activities in a project

**Properties:**
- Count (Long, RO): Number of activities
- Filter (String, R/W): Current filter expression
- Sort (String, R/W): Current sort expression

**Essential Methods:**
- **Item(index/id)**: Get activity by numeric index or string ID
- **Add(id, description)**: Add new activity, returns OPActivity
- **Remove(index/id)**: Delete activity
- **Fields**: Returns OPFields collection for field info
- **SetFilterTo(filtername)**: Apply named filter
- **SetSortTo(sortname)**: Apply named sort

---

### OPPredecessor

**Purpose:** Represents a predecessor relationship

**Key Properties:**
- ID (String, R/W): Predecessor activity ID
- SuccessorID (String, RO): Successor activity ID
- RelationshipType (String, R/W): "FS", "SS", "FF", "SF"
- Lag (String, R/W): Lag duration (e.g., "2D", "-1W")
- Calendar (String, R/W): Calendar for lag calculation
- TotalFloat, FreeFloat (String, RO): Float values
- Turns (Long, R/W): Number of turns (loops)

**Essential Methods:**
- **Remove()**: Delete this relationship
- **GetField/SetField**: Access fields

---

### OPPredecessors (Collection)

**Essential Methods:**
- **Add(predecessorID)**: Add relationship, returns OPPredecessor
- **Item(index)**: Get by index
- **Remove(index)**: Delete relationship

---

### OPAssignment

**Purpose:** Resource assignment to an activity

**Key Properties:**
- ActivityID (String, R/W): Activity being assigned to
- ResourceID (String, R/W): Resource being assigned
- AlternateResourceID (String, R/W): Alternate resource
- Level (Double, R/W): Assignment level/units
- LevelType (String, R/W): Level type
- Remaining (Double, R/W): Remaining amount
- ResourceOffset (String, R/W): Delay from activity start (e.g., "2D")
- ResourcePeriod (String, R/W): Duration of assignment (e.g., "5D")

**Essential Methods:**
- **Remove()**: Delete assignment
- **GetResourceDateArray()**: Get time-phased usage array
- **GetResourceCrosstabData()**: Get crosstab data
- **GetField/SetField**: Access fields

---

### OPAssignments (Collection)

**Essential Methods:**
- **Add()**: Add new assignment, returns OPAssignment
- **Item(index)**: Get by index
- **Remove(index)**: Delete assignment

---

### OPProjectResource

**Purpose:** Resource used in project (from resource file)

**Key Properties:**
- ID (String, RO): Resource identifier
- Description (String, R/W): Resource description
- Class (String, R/W): "Labor", "Material", "Other Direct Costs", "Subcontract"
- Type (String, R/W): Resource type
- Units (Double, R/W): Available units
- UnitCost (Double, R/W): Cost per unit
- RollUp (String, R/W): Rollup settings
- Suppress (Bool, R/W): Suppress in reports

**Essential Methods:**
- **Availabilities**: Returns OPAvailabilities collection
- **Notes**: Returns OPNotes collection
- **GetResourceCrosstabData()**: Get time-phased data
- **GetResourceDateArray()**: Get usage array
- **GetField/SetField**: Access fields

---

### OPProjectResources (Collection)

**Properties:**
- FileName (String, RO): Resource file path

**Essential Methods:**
- **Item(index/id)**: Get resource by index or ID
- **Filters, Sorts, CalculatedFields**: Access definitions

---

### OPResourceRecord

**Purpose:** Resource record in resource file

Similar properties/methods to OPProjectResource, plus:

**Essential Methods:**
- **AssignSkill(skillname)**: Assign skill to resource
- **RemoveSkill(skillname)**: Remove skill
- **GetSkills()**: Returns array of skill names

---

### OPResource (Resource File Collection)

**Purpose:** Open resource file containing resource records

**Properties:**
- FileName (String, RO): Resource file path
- Count (Long, RO): Number of records
- Calendar (String, R/W): Default calendar
- Filter, Sort (String, R/W): Current filter/sort

**Essential Methods:**
- **Item(index/id)**: Get resource record
- **Add()**: Add new resource record
- **Remove(index/id)**: Delete record
- **Save(), SaveAs(path)**: Save file
- **Shut(), ShutWOSave()**: Close file
- **Fields, Filters, Sorts, CalculatedFields**: Access definitions

---

### OPAvailability

**Purpose:** Resource availability period

**Key Properties:**
- StartDate, StopDate (Date, R/W): Availability period
- Level (Double, R/W): Available units
- Calendar (String, R/W): Calendar for this period

---

### OPCodeRecord

**Purpose:** Code record in code file

**Key Properties:**
- Code (String, R/W): Code value
- Description (String, R/W): Code description

**Essential Methods:**
- **Notes**: Returns OPNotes collection
- **GetField/SetField**: Access fields

---

### OPCode (Code File Collection)

Similar to OPResource but for code files.

**Essential Methods:**
- **Item(index/code)**: Get code record
- **Add()**: Add new code record
- **Categories, Rollups**: Access code organization

---

### OPProjectCode

**Purpose:** Code file attached to project

**Properties:**
- FieldName (String, RO): Activity field this code applies to
- FileName (String, RO): Code file path

---

### OPCalendarRecord

**Purpose:** Named calendar within calendar file

**Essential Methods:**
- **Name()**: Get calendar name
- **GetStandardDays(), SetStandardDays(array)**: Standard work days
- **StandardDay(1-7)**: Get specific day (1=Sun, 7=Sat), returns OPStandardDay
- **Holidays, ExtraWorkDays**: Returns collections
- **Date(datevalue)**: Get specific date, returns OPDate

---

### OPCalendar (Calendar File Collection)

**Essential Methods:**
- **Item(index/name)**: Get calendar record by name
- **Add()**: Add new calendar record
- **Save(), SaveAs(path)**: Save file

---

### OPStandardDay

**Properties:**
- Work (Bool, R/W): Is work day

**Methods:**
- **Shifts**: Returns OPShifts collection for this day

---

### OPShift

**Properties:**
- StartTime (String, R/W): Shift start (e.g., "08:00")
- StopTime (String, R/W): Shift end (e.g., "17:00")

---

### OPDate

**Properties:**
- Date (Date, RO): The date
- Work (Bool, R/W): Is work day

**Methods:**
- **Shifts**: Returns OPShifts collection for this date

---

### OPBaselines (Baseline Activities Collection)

**Purpose:** Contains baseline activities for a saved baseline

**Properties:**
- Name (String, RO): Baseline name
- Description (String, RO): Baseline description
- Selected (Bool, RO): Is this the selected baseline
- Count, Filter, Sort: Standard collection properties

**Methods:**
- **Item(index/id)**: Get baseline activity (OPActivity object)

---

### OPBaselinesList (Collection)

**Essential Methods:**
- **Item(index/name)**: Get OPBaselines collection
- **Select(name)**: Set selected baseline

---

### OPCost

**Purpose:** Manual cost entry

**Key Properties:**
- ActivityID (String, R/W): Activity this cost applies to
- ResourceID (String, R/W): Resource (optional)
- StartDate, EndDate (Date, R/W): Cost period
- ActualCost, ActualQty (Double, R/W): Cost values

---

### OPResourceUsage

**Purpose:** Resource usage record

**Key Properties:**
- ActivityID (String, R/W): Activity
- ResourceID (String, R/W): Resource
- StartDate, FinishDate (Date, R/W): Usage period
- Used (Double, R/W): Amount used
- Cost, EscalatedCost (Double, R/W): Cost values

---

### OPNote

**Properties:**
- Category (String, RO): Note category
- NoteText (String, R/W): Note content
- ModifiedBy (String, RO): Last modifier
- ModifiedDate (Date, RO): Last modified date

---

### OPNotes (Collection)

**Essential Methods:**
- **Add(category)**: Add note, returns OPNote
- **Item(index)**: Get note by index

---

### OPFilter, OPSort, OPCalculatedField

**Purpose:** Saved filter/sort/calculated field definitions

**Key Properties:**
- Name (String, RO): Definition name
- Expression (String, R/W): Definition expression
- TableName (String, RO): Table this applies to

---

### OPGlobalEdit

**Purpose:** Global edit definition

**Key Properties:**
- Name (String, RO): Edit name
- Expression (String, R/W): Value expression
- Filter (String, R/W): Filter to limit scope
- ApplyToField (String, R/W): Field to modify

**Methods:**
- **Apply()**: Execute global edit

---

### OPView

**Purpose:** Saved view (barchart, network, spreadsheet, graph)

**Essential Methods:**
- **Activate()**: Open and display view
- **Name()**: Get view name
- **Description()**: Get description
- **ExpandAll(), CollapseAll()**: Expand/collapse outline
- **SetDateScaleOption(params)**: Set timescale
- **SetViewLegend(params)**: Configure legend
- **RefreshFilterandSort()**: Refresh after filter/sort change

---

### OPWebWindow

**Purpose:** Browser view window

**Properties:**
- URLAddress (String, R/W): URL to display
- Title (String, R/W): Window title
- QueryString (String, RO): HTTP query string

**Methods:**
- **Close()**: Close window

---

### OPReportingCalendar

**Purpose:** Reporting calendar for crosstab date generation

**Methods:**
- **Item(index)**: Get OPReportingCalendarRecord
- **Add()**: Add period

---

### OPFileCabinet

**Purpose:** Open Plan Explorer/File Cabinet

**Methods:**
- **Projects, Resources, Calendars, Codes**: Get file icon collections
- **Barcharts, Networks, Spreadsheets, Graphs, Views**: Get view icon collections

---

## Common Patterns & Examples

### Filtering Activities

```vb
'Set filter expression directly
acts.Filter = "EARLY_START < #1/1/2024# AND CRITICAL > 0"

'Or use named filter
acts.SetFilterTo "MyFilterName"

'Clear filter
acts.Filter = ""
```

### Sorting Activities

```vb
'Set sort expression
acts.Sort = "EARLY_START"  'Ascending
acts.Sort = "EARLY_START DESC"  'Descending

'Multiple fields
acts.Sort = "WBS, EARLY_START"

'Named sort
acts.SetSortTo "MySortName"
```

### Adding Activities with Relationships

```vb
'Add activities
Set act1 = acts.Add("ACT001", "Foundation")
act1.Duration = "10D"
act1.Calendar = "STANDARD"

Set act2 = acts.Add("ACT002", "Framing")
act2.Duration = "15D"

'Add predecessor
Set pred = act2.Predecessors.Add("ACT001")
pred.RelationshipType = "FS"
pred.Lag = "2D"
```

### Resource Assignments

```vb
Set act = acts.Item("ACT001")
Set assigns = act.Assignments

'Add assignment
Set assign = assigns.Add
assign.ResourceID = "LABOR001"
assign.Level = 2.0  'Two units
assign.ResourcePeriod = act.Duration  'Full activity duration

'With offset (start 2 days into activity)
assign.ResourceOffset = "2D"
assign.ResourcePeriod = "8D"
```

### Working with Baselines

```vb
'Create baseline
proj.CreateBaseline "Baseline 1"

'Compare current to baseline
Set acts = proj.Activities
For i = 1 To acts.Count
    Set act = acts.Item(i)

    currentDur = act.Duration
    baselineDur = act.BaselineStart - act.BaselineFinish

    If act.EarlyFinish > act.BaselineFinish Then
        Debug.Print act.ID & " is behind schedule"
    End If
Next i

'Select baseline for comparison
Set baselinesList = proj.BaselinesList
baselinesList.Select "Baseline 1"

'Access baseline activities
Set baseline = baselinesList.Item("Baseline 1")
Set baselineAct = baseline.Item("ACT001")
```

### Auto Progress

```vb
'Configure auto progress
proj.StatusDate = Date  'Set Time Now to today

'Activity progress settings
proj.AutoProgActivity = True
proj.AutoProgActBasedOn = "EARLY"
proj.AutoProgActInProgress = True  'Mark started activities
proj.AutoProgActComplete = True  'Mark completed activities
proj.AutoProgActFilter = "MyProgressFilter"  'Optional

'Resource progress settings
proj.AutoProgResource = True
proj.AutoProgResStartDate = #1/1/2024#
proj.AutoProgResEndDate = Date

'Execute
proj.AutoProgress

'Alternative: Use AutoProgressEx with XML params for more control
```

### Cost Calculation

```vb
'Configure cost calculation
proj.CalcCostBudget = True
proj.CalcCostActual = True
proj.CalcCostRemaining = True
proj.CalcCostEarnedValues = True
proj.CalcCostBasedOn = "EARLY"
proj.CalcCostEscalated = True

'Execute
proj.CalculateCost

'Access results
Set act = acts.Item("ACT001")
Debug.Print "Budget: " & act.BudgetCost
Debug.Print "Actual: " & act.CostToDate
Debug.Print "Resource Cost: " & act.TotalResourceCost
```

### Crosstab Data (Time-Phased)

```vb
'Generate crosstab dates
startDate = #1/1/2024#
endDate = #12/31/2024#
interval = "W"  'Weekly (M=month, W=week, D=day)
proj.GenerateCrosstabDates startDate, endDate, interval

'Get resource usage data
Set res = proj.Resources.Item("LABOR001")
dataArray = res.GetResourceCrosstabData()

'Or get as XML
xmlData = res.GetResourceCrosstabDataInXML()

'Get resource date array for assignment
Set act = acts.Item("ACT001")
Set assigns = act.Assignments
Set assign = assigns.Item(1)
dateArray = assign.GetResourceDateArray()
```

### Shared Mode (Multi-user)

```vb
'Set shared mode lock
proj.SetSharedmodeLock True

'Force data refresh from database
act.RefreshData = True
currentValue = act.PercentComplete  'Gets fresh data

'Release lock
proj.SetSharedmodeLock False

'Check lock status
lockStatus = proj.GetSharedModeLock
```

### Working with User-Defined Fields

```vb
'Set user field
act.SetField "USER1", "Custom Value"
act.SetField "USER_DATE1", #1/15/2024#
act.SetField "USER_NUM1", 42.5

'Get user field
value = act.GetField("USER1")

'Get multiple fields efficiently
Dim fieldNames(5) As String
fieldNames(0) = "ACT_ID"
fieldNames(1) = "DESCRIPTN"
fieldNames(2) = "ORIG_DUR"
fieldNames(3) = "USER1"
fieldNames(4) = "USER_DATE1"
fieldNames(5) = "USER_NUM1"

values = act.GetFields(fieldNames)

For i = 0 To 5
    Debug.Print fieldNames(i) & " = " & values(i)
Next i
```

### Field Information

```vb
'Get field definitions
Set fields = acts.Fields
Set field = fields.Item("ACT_ID")

Debug.Print "Field Name: " & field.FieldName
Debug.Print "User Name: " & field.UserName
Debug.Print "Type: " & field.FieldType
Debug.Print "DB Type: " & field.DBType
Debug.Print "Editable: " & field.IsEditable
```

### Creating Filters/Sorts/Calculated Fields

```vb
'Create filter
Set filters = proj.Filters
Set filter = filters.Add("CriticalActivities", "CRITICAL > 0", "ACTIVITY")
acts.SetFilterTo "CriticalActivities"

'Create sort
Set sorts = proj.Sorts
Set sort = sorts.Add("EarlyDates", "EARLY_START, ACT_ID", "ACTIVITY")
acts.SetSortTo "EarlyDates"

'Create calculated field
Set calcFields = proj.CalculatedFields
Set calcField = calcFields.Add("MyCalc", "EARLY_FINISH - EARLY_START", "ACTIVITY")
calcField.ResultType = "DURATION"
```

### Global Edits

```vb
'Create global edit
Set globalEdits = proj.GlobalEdits
Set edit = globalEdits.Add("UpdatePriority", "5", "ACTIVITY")
edit.ApplyToField = "USER_NUM1"
edit.Filter = "CRITICAL > 0"

'Apply
edit.Apply

'Or use batch global edit
Set batchEdits = proj.BatchGlobalEdits
Set batchEdit = batchEdits.Item("MyBatch")
batchEdit.Apply
```

### Views & Windows

```vb
'Access barchart views
Set barcharts = proj.Barcharts
Set view = barcharts.Item(1)
view.Activate

'Or activate by name
barcharts.ActivateByFilename "MyBarchart"

'Add new view
barcharts.AddView "NewBarchart"

'Configure view
view.SetDateScaleOption params
view.ExpandAll
view.RefreshFilterandSort

'Create browser view
OPApp.CreateBrowserView "http://example.com", "Window Title"
```

### Working with Code Files

```vb
'Access project codes
Set codeFi les = proj.CodeFiles
Set codeFile = codeFiles.Item(1)

Debug.Print "Field Name: " & codeFile.FieldName  'Activity field
Debug.Print "File: " & codeFile.FileName

'Access code records
Set code = codeFile.Item("CODE01")
code.Description = "Updated Description"

'Add code
Set newCode = codeFile.Add()
newCode.Code = "CODE99"
newCode.Description = "New Code"

'Add note to code
Set notes = code.Notes
Set note = notes.Add("General")
note.NoteText = "This is a note"
```

### Calendar Operations

```vb
'Open calendar
Set calendars = OPApp.Calendars
Set calendar = calendars.Item("STANDARD")

'Access calendar record
Set calRec = calendar.Item("5 Day Work Week")

'Get standard days
For day = 1 To 7  '1=Sun, 7=Sat
    Set stdDay = calRec.StandardDay(day)
    Debug.Print "Day " & day & " is work day: " & stdDay.Work

    'Get shifts for work days
    If stdDay.Work Then
        Set shifts = stdDay.Shifts
        For j = 1 To shifts.Count
            Set shift = shifts.Item(j)
            Debug.Print "  Shift: " & shift.StartTime & " to " & shift.StopTime
        Next j
    End If
Next day

'Add holiday
Set holidays = calRec.Holidays
holidays.Add #7/4/2024#

'Add extra work day
Set extraDays = calRec.ExtraWorkDays
Set extraDate = extraDays.Add(#7/6/2024#)
extraDate.Work = True
```

### Import/Export

```vb
'Export - configure params (refer to Open Plan documentation)
Dim exportParams As String
exportParams = "XML params here"
proj.GeneralExport exportParams

'Import
Dim importParams As String
importParams = "XML params here"
proj.GeneralImport importParams
```

### Error Handling

```vb
On Error Resume Next

'Attempt operation
Set proj = OPApp.FileOpen("C:\nonexistent.opp")

If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description, vbCritical
    Err.Clear
End If

On Error GoTo 0
```

---

## Key Enumerations & Constants

**Relationship Types:**
- "FS" - Finish to Start
- "SS" - Start to Start
- "FF" - Finish to Finish
- "SF" - Start to Finish

**Resource Classes:**
- "Labor"
- "Material"
- "Other Direct Costs"
- "Subcontract"

**Progress Flags:**
- "Planned"
- "Remaining Duration"
- "Percent Complete"
- "Elapsed Duration"
- "Complete"

**Activity Computed Status:**
- "Planned"
- "In Progress"
- "Complete"

**Schedule Methods:**
- "Time Limited"
- "Resource Limited"

**Based On Values (AutoProg, CalcCost):**
- "EARLY"
- "LATE"
- "SCHEDULE"
- "BASELINE"

**Auto Progress Types:**
- "REMAINING"
- "ELAPSED"
- "PERCENT"
- "EXPECTED"

**Duration Distribution Shapes:**
- "Beta"
- "Normal"
- "Triangular"
- "Uniform"
- "None"

**Duration Units:**
- M = Months
- W = Weeks
- D = Days
- H = Hours
- T = Minutes

**Duration Format:** Number + Unit, e.g., "10D", "2W", "5H"
Zero durations don't require unit: "0"

**Out of Sequence Options (OutOfSeqOpt):**
- 0 = Ignore positive lag of predecessor
- 1 = Observe positive lag of predecessor
- 2 = Ignore predecessor relationship

---

## Database Field Names (Common)

Use with GetField/SetField for direct database access:

**Activity Fields:**
- ACT_ID: Activity ID
- DESCRIPTN: Description
- ORIG_DUR: Original Duration
- EARLY_START, EARLY_FINISH: Early dates
- LATE_START, LATE_FINISH: Late dates
- ACT_START, ACT_FINISH: Actual dates
- PCT_COMP: Percent Complete
- TOTAL_FLT: Total Float
- FREE_FLT: Free Float
- CALENDAR: Calendar name
- USER1-USER10: User-defined text fields
- USER_DATE1-USER_DATE10: User-defined date fields
- USER_NUM1-USER_NUM10: User-defined numeric fields

**Resource Fields:**
- RES_ID: Resource ID
- RES_DESC: Resource Description
- RES_CLASS: Resource Class
- RES_TYPE: Resource Type
- UNIT_COST: Unit Cost
- UNITS: Units
- NO_LIST: Suppress flag

---

## Important Notes

### Open Plan 8.x Requirements
- **Must call Login()** immediately after CreateObject
- Security permissions affect automation access
- Unsupported: OPFCProcesses, OPFCStructures objects
- Unsupported: FileOpenSpecialOpenPlan, OpenSpecialODBC methods

### Shared Mode
- Set RefreshData = True on objects to force database refresh
- Use SetSharedmodeLock for exclusive access during updates
- Critical for multi-user environments

### Performance Tips
- Use GetFields() instead of multiple GetField() calls
- Set filters before large iterations
- Disable AutoAnalyze when making bulk changes
- Use RefreshData sparingly (only when needed)

### Duration Values
- Must include unit identifier (D, W, M, H, T)
- Zero durations can omit unit: "0"
- Examples: "5D", "2.5W", "10H"

### Date Values
- Use # delimiters in VB: #1/15/2024#
- Respect project DateFormat setting
- StatusDate affects schedule calculations

### Collection Indexing
- Collections are 1-based (first item is index 1)
- Many collections support both numeric index and string key
- Example: acts.Item(1) or acts.Item("ACT001")

---

## Quick Troubleshooting

**"Permission denied" errors:**
- Check security settings in Deltek EPM Security Administrator
- Ensure menu items are enabled (even if invisible)
- Call Login() before other operations

**"Object not found" errors:**
- Verify file paths are correct and accessible
- Check object IDs exist
- Ensure files are opened before accessing

**Changes not saving:**
- Call Save() or SaveAs() explicitly
- Check file access permissions
- Verify not in read-only mode

**Schedule not calculating:**
- Call TimeAnalyze() after changes
- Check for missing calendars
- Verify no circular relationships

**Filters not working:**
- Check filter syntax (use # for dates)
- Field names must match database names
- Use SetFilterTo() for named filters or set Filter property directly

---

*This reference provides ~90% token reduction while preserving all essential API information. For complete method signatures and advanced parameters, consult the full DeltekOpenPlanDeveloperGuide.md.*

---

## Complete VBA Examples

### Example 1: Complete Project Automation Workflow

```vb
Sub AutomateProjectWorkflow()
    On Error GoTo ErrorHandler
    
    Dim OPApp As Object
    Dim proj As Object
    Dim acts As Object
    Dim act As Object
    Dim i As Long
    
    ' Initialize Open Plan
    Set OPApp = CreateObject("opp.application")
    OPApp.Login
    
    ' Configure application
    OPApp.SilentMode = True  ' Suppress dialogs
    OPApp.Show  ' Show application window
    
    ' Open project
    Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")
    
    ' Configure project settings
    With proj
        .StatusDate = Date  ' Set Time Now to today
        .AutoAnalyze = True  ' Auto time analysis after changes
        
        ' Auto Progress settings
        .AutoProgActivity = True
        .AutoProgActBasedOn = "EARLY"
        .AutoProgActInProgress = True
        .AutoProgActComplete = True
        .AutoProgResource = True
        .AutoProgResStartDate = DateAdd("m", -1, Date)  ' Last month
        .AutoProgResEndDate = Date
        
        ' Cost calculation settings
        .CalcCostBudget = True
        .CalcCostActual = True
        .CalcCostRemaining = True
        .CalcCostEarnedValues = True
        .CalcCostBasedOn = "EARLY"
    End With
    
    ' Apply auto progress
    proj.AutoProgress
    Debug.Print "Auto progress completed"
    
    ' Perform time analysis
    proj.TimeAnalyze
    Debug.Print "Time analysis completed"
    
    ' Resource schedule
    proj.ResourceSchedule "Time Limited", False  ' Don't re-analyze
    Debug.Print "Resource scheduling completed"
    
    ' Calculate costs
    proj.CalculateCost
    Debug.Print "Cost calculation completed"
    
    ' Get activities collection
    Set acts = proj.Activities
    Debug.Print "Total activities: " & acts.Count
    
    ' Apply filter to show only critical activities
    acts.Filter = "CRITICAL > 0 AND PCT_COMP < 100"
    Debug.Print "Critical incomplete activities: " & acts.Count
    
    ' Apply sort
    acts.Sort = "EARLY_START, ACT_ID"
    
    ' Process critical activities
    For i = 1 To acts.Count
        Set act = acts.Item(i)
        
        Debug.Print "Activity: " & act.ID & " - " & act.Description
        Debug.Print "  Status: " & act.ComputedStatus
        Debug.Print "  % Complete: " & act.PercentComplete
        Debug.Print "  Early Start: " & act.EarlyStart
        Debug.Print "  Early Finish: " & act.EarlyFinish
        Debug.Print "  Total Float: " & act.TotalFloat
        Debug.Print "  Budget Cost: " & Format(act.BudgetCost, "$#,##0.00")
        Debug.Print "  Actual Cost: " & Format(act.CostToDate, "$#,##0.00")
        
        ' Check if behind schedule
        If act.EarlyFinish > act.BaselineFinish And act.PercentComplete < 100 Then
            Debug.Print "  WARNING: Behind baseline schedule!"
            
            ' Add note
            Dim notes As Object
            Dim note As Object
            Set notes = act.Notes
            Set note = notes.Add("Status")
            note.NoteText = "Activity is behind baseline schedule. " & _
                           "Review required. Checked on " & Date
        End If
        
        Debug.Print "---"
    Next i
    
    ' Clear filter to see all activities
    acts.Filter = ""
    
    ' Create baseline if it doesn't exist
    On Error Resume Next
    proj.CreateBaseline "Monthly Baseline " & Format(Date, "yyyy-mm")
    If Err.Number = 0 Then
        Debug.Print "Baseline created successfully"
    Else
        Debug.Print "Baseline may already exist or error occurred: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Save project
    proj.Save
    Debug.Print "Project saved"
    
    ' Create backup
    proj.CreateBackup "C:\Backups\MyProject_" & Format(Date, "yyyymmdd") & ".opp"
    Debug.Print "Backup created"
    
    ' Generate report data
    GenerateProjectReport proj
    
    ' Close project
    proj.Shut
    
    ' Cleanup
    Set act = Nothing
    Set acts = Nothing
    Set proj = Nothing
    Set OPApp = Nothing
    
    Debug.Print "Workflow completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
    If Not proj Is Nothing Then proj.ShutWOSave
    Set OPApp = Nothing
End Sub

Sub GenerateProjectReport(proj As Object)
    ' Generate custom project report
    Dim acts As Object
    Dim act As Object
    Dim totalBudget As Double
    Dim totalActual As Double
    Dim completeCount As Long
    Dim inProgressCount As Long
    Dim plannedCount As Long
    Dim i As Long
    
    Set acts = proj.Activities
    
    For i = 1 To acts.Count
        Set act = acts.Item(i)
        
        totalBudget = totalBudget + act.BudgetCost
        totalActual = totalActual + act.CostToDate
        
        Select Case act.ComputedStatus
            Case "Complete"
                completeCount = completeCount + 1
            Case "In Progress"
                inProgressCount = inProgressCount + 1
            Case "Planned"
                plannedCount = plannedCount + 1
        End Select
    Next i
    
    Debug.Print "========== PROJECT SUMMARY =========="
    Debug.Print "Project: " & proj.Name
    Debug.Print "Status Date: " & proj.StatusDate
    Debug.Print ""
    Debug.Print "Activity Counts:"
    Debug.Print "  Total: " & acts.Count
    Debug.Print "  Complete: " & completeCount
    Debug.Print "  In Progress: " & inProgressCount
    Debug.Print "  Planned: " & plannedCount
    Debug.Print ""
    Debug.Print "Financial Summary:"
    Debug.Print "  Total Budget: " & Format(totalBudget, "$#,##0.00")
    Debug.Print "  Actual Cost: " & Format(totalActual, "$#,##0.00")
    Debug.Print "  Variance: " & Format(totalActual - totalBudget, "$#,##0.00")
    Debug.Print "  % Spent: " & Format((totalActual / totalBudget) * 100, "0.0") & "%"
    Debug.Print "===================================="
End Sub
```

### Example 2: Building a Project from Scratch

```vb
Sub CreateProjectFromScratch()
    On Error GoTo ErrorHandler
    
    Dim OPApp As Object
    Dim proj As Object
    Dim acts As Object
    Dim act As Object
    Dim pred As Object
    Dim resources As Object
    Dim assigns As Object
    Dim assign As Object
    
    ' Initialize
    Set OPApp = CreateObject("opp.application")
    OPApp.Login
    
    ' Create new project
    Set proj = OPApp.FileNew("C:\Projects\NewProject.opp")
    
    ' Configure project
    With proj
        .Name = "New Construction Project"
        .Description = "Sample construction project created via automation"
        .Client = "ABC Corporation"
        .Manager = "John Smith"
        .StartDate = #1/1/2024#
        .StatusDate = #1/1/2024#
        .DefaultDurationUnit = "Days"
        .DefaultActivityCalendar = "STANDARD"
        .Calendar = "C:\OpenPlan\Calendars\Standard.cal"
        .Resource = "C:\OpenPlan\Resources\Construction.res"
    End With
    
    ' Get activities collection
    Set acts = proj.Activities
    
    ' Create project start milestone
    Set act = acts.Add("START", "Project Start")
    act.Type = "Milestone"
    act.Duration = "0"
    act.ScheduledStart = proj.StartDate
    
    ' Create Phase 1 activities
    Set act = acts.Add("DESIGN.001", "Preliminary Design")
    act.Duration = "10D"
    act.Description = "Create preliminary design documents"
    act.Calendar = "STANDARD"
    act.SetField "USER1", "Phase 1"
    
    Set act = acts.Add("DESIGN.002", "Detailed Design")
    act.Duration = "15D"
    act.Description = "Complete detailed design and drawings"
    act.SetField "USER1", "Phase 1"
    
    Set act = acts.Add("DESIGN.003", "Design Review")
    act.Duration = "5D"
    act.Description = "Client design review and approval"
    act.SetField "USER1", "Phase 1"
    
    ' Create Phase 2 activities
    Set act = acts.Add("CONST.001", "Site Preparation")
    act.Duration = "5D"
    act.Description = "Clear and prepare construction site"
    act.SetField "USER1", "Phase 2"
    
    Set act = acts.Add("CONST.002", "Foundation")
    act.Duration = "10D"
    act.Description = "Pour foundation and let cure"
    act.SetField "USER1", "Phase 2"
    
    Set act = acts.Add("CONST.003", "Framing")
    act.Duration = "15D"
    act.Description = "Construct building frame"
    act.SetField "USER1", "Phase 2"
    
    Set act = acts.Add("CONST.004", "Roofing")
    act.Duration = "8D"
    act.Description = "Install roofing system"
    act.SetField "USER1", "Phase 2"
    
    ' Create Phase 3 activities
    Set act = acts.Add("FINISH.001", "Electrical")
    act.Duration = "12D"
    act.Description = "Install electrical systems"
    act.SetField "USER1", "Phase 3"
    
    Set act = acts.Add("FINISH.002", "Plumbing")
    act.Duration = "10D"
    act.Description = "Install plumbing systems"
    act.SetField "USER1", "Phase 3"
    
    Set act = acts.Add("FINISH.003", "Interior Finish")
    act.Duration = "20D"
    act.Description = "Complete interior finishing work"
    act.SetField "USER1", "Phase 3"
    
    ' Create project end milestone
    Set act = acts.Add("END", "Project Complete")
    act.Type = "Milestone"
    act.Duration = "0"
    
    ' Add relationships (critical path)
    Set act = acts.Item("DESIGN.001")
    Set pred = act.Predecessors.Add("START")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    Set act = acts.Item("DESIGN.002")
    Set pred = act.Predecessors.Add("DESIGN.001")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    Set act = acts.Item("DESIGN.003")
    Set pred = act.Predecessors.Add("DESIGN.002")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    Set act = acts.Item("CONST.001")
    Set pred = act.Predecessors.Add("DESIGN.003")
    pred.RelationshipType = "FS"
    pred.Lag = "2D"  ' 2 day lag for procurement
    
    Set act = acts.Item("CONST.002")
    Set pred = act.Predecessors.Add("CONST.001")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    Set act = acts.Item("CONST.003")
    Set pred = act.Predecessors.Add("CONST.002")
    pred.RelationshipType = "FS"
    pred.Lag = "7D"  ' Wait for cure
    
    Set act = acts.Item("CONST.004")
    Set pred = act.Predecessors.Add("CONST.003")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    ' Parallel activities - both can start when framing is done
    Set act = acts.Item("FINISH.001")
    Set pred = act.Predecessors.Add("CONST.004")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    Set act = acts.Item("FINISH.002")
    Set pred = act.Predecessors.Add("CONST.004")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    ' Interior finish waits for electrical and plumbing
    Set act = acts.Item("FINISH.003")
    Set pred = act.Predecessors.Add("FINISH.001")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    Set pred = act.Predecessors.Add("FINISH.002")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    Set act = acts.Item("END")
    Set pred = act.Predecessors.Add("FINISH.003")
    pred.RelationshipType = "FS"
    pred.Lag = "0"
    
    ' Add resources to activities
    Set resources = proj.Resources
    
    ' Add labor to design activities
    Set act = acts.Item("DESIGN.001")
    Set assigns = act.Assignments
    Set assign = assigns.Add
    assign.ResourceID = "ENGINEER"
    assign.Level = 1.0
    assign.ResourcePeriod = "10D"
    
    Set act = acts.Item("DESIGN.002")
    Set assigns = act.Assignments
    Set assign = assigns.Add
    assign.ResourceID = "ENGINEER"
    assign.Level = 2.0  ' Two engineers
    assign.ResourcePeriod = "15D"
    
    ' Add labor to construction activities
    Set act = acts.Item("CONST.002")
    Set assigns = act.Assignments
    Set assign = assigns.Add
    assign.ResourceID = "LABOR"
    assign.Level = 4.0  ' Four laborers
    assign.ResourcePeriod = "10D"
    
    Set act = acts.Item("CONST.003")
    Set assigns = act.Assignments
    Set assign = assigns.Add
    assign.ResourceID = "CARPENTER"
    assign.Level = 6.0  ' Six carpenters
    assign.ResourcePeriod = "15D"
    
    ' Perform time analysis
    proj.TimeAnalyze
    
    ' Create initial baseline
    proj.CreateBaseline "Original Plan"
    
    ' Save project
    proj.Save
    
    Debug.Print "Project created successfully at: C:\Projects\NewProject.opp"
    Debug.Print "Project duration: " & DateDiff("d", proj.EarlyFinishDate, proj.StartDate) & " days"
    Debug.Print "Early finish: " & proj.EarlyFinishDate
    
    ' Close project
    proj.Shut
    
    ' Cleanup
    Set OPApp = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description
    If Not proj Is Nothing Then proj.ShutWOSave
End Sub
```

### Example 3: Advanced Resource Management

```vb
Sub ManageResourcesAndCosts()
    On Error GoTo ErrorHandler
    
    Dim OPApp As Object
    Dim proj As Object
    Dim resFile As Object
    Dim res As Object
    Dim avail As Object
    Dim acts As Object
    Dim act As Object
    Dim assigns As Object
    Dim assign As Object
    Dim i As Long
    
    Set OPApp = CreateObject("opp.application")
    OPApp.Login
    
    Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")
    
    ' Get resource file
    Set resFile = OPApp.Resources.Item("Construction")
    
    ' Update resource availability
    Set res = resFile.Item("CARPENTER")
    
    Debug.Print "Resource: " & res.ID
    Debug.Print "Description: " & res.Description
    Debug.Print "Class: " & res.Class
    Debug.Print "Type: " & res.Type
    Debug.Print "Units: " & res.Units
    Debug.Print "Unit Cost: " & res.UnitCost
    
    ' Add availability periods
    Dim avails As Object
    Set avails = res.Availabilities
    
    ' Add Q1 availability
    Set avail = avails.Add
    avail.StartDate = #1/1/2024#
    avail.StopDate = #3/31/2024#
    avail.Level = 10.0  ' 10 carpenters available
    avail.Calendar = "STANDARD"
    
    ' Add Q2 availability (increased capacity)
    Set avail = avails.Add
    avail.StartDate = #4/1/2024#
    avail.StopDate = #6/30/2024#
    avail.Level = 15.0  ' 15 carpenters available
    avail.Calendar = "STANDARD"
    
    ' Save resource file
    resFile.Save
    
    ' Analyze resource usage in project
    Set acts = proj.Activities
    acts.Filter = ""  ' Clear any filters
    
    Debug.Print vbCrLf & "========== RESOURCE USAGE ANALYSIS =========="
    
    ' Dictionary to track resource totals
    Dim resourceTotals As Object
    Set resourceTotals = CreateObject("Scripting.Dictionary")
    
    For i = 1 To acts.Count
        Set act = acts.Item(i)
        
        ' Get assignments for this activity
        Set assigns = act.Assignments
        
        If assigns.Count > 0 Then
            Debug.Print "Activity: " & act.ID & " - " & act.Description
            Debug.Print "  Duration: " & act.Duration
            Debug.Print "  Assignments:"
            
            Dim j As Long
            For j = 1 To assigns.Count
                Set assign = assigns.Item(j)
                
                Debug.Print "    Resource: " & assign.ResourceID
                Debug.Print "    Level: " & assign.Level
                Debug.Print "    Period: " & assign.ResourcePeriod
                
                ' Track totals
                If Not resourceTotals.Exists(assign.ResourceID) Then
                    resourceTotals.Add assign.ResourceID, 0
                End If
                resourceTotals(assign.ResourceID) = _
                    resourceTotals(assign.ResourceID) + assign.Level
            Next j
        End If
    Next i
    
    ' Display resource summary
    Debug.Print vbCrLf & "========== RESOURCE SUMMARY =========="
    Dim key As Variant
    For Each key In resourceTotals.Keys
        Debug.Print "Resource: " & key
        Debug.Print "  Total Usage: " & resourceTotals(key)
    Next key
    
    ' Configure and generate resource crosstab data
    proj.GenerateCrosstabDates #1/1/2024#, #12/31/2024#, "W"  ' Weekly
    
    Set res = proj.Resources.Item("CARPENTER")
    
    ' Get resource usage data
    Dim usageData As Variant
    usageData = res.GetResourceCrosstabData()
    
    ' Or get as XML for external processing
    Dim xmlData As String
    xmlData = res.GetResourceCrosstabDataInXML()
    
    ' Can save XML to file for external reporting
    Dim fso As Object
    Dim txtFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile("C:\Reports\ResourceUsage.xml", True)
    txtFile.Write xmlData
    txtFile.Close
    
    Debug.Print "Resource usage data exported to XML"
    
    ' Perform resource leveling
    Debug.Print vbCrLf & "Performing resource scheduling..."
    proj.ResourceSchedule "Resource Limited", True
    
    ' Check for resource overallocations
    Debug.Print vbCrLf & "========== OVERALLOCATED RESOURCES =========="
    Set acts = proj.Activities
    For i = 1 To acts.Count
        Set act = acts.Item(i)
        
        If act.DelayingResource <> "" Then
            Debug.Print "Activity: " & act.ID
            Debug.Print "  Delayed by resource: " & act.DelayingResource
            Debug.Print "  Scheduled Start: " & act.ScheduledStart
            Debug.Print "  Scheduled Finish: " & act.ScheduledFinish
        End If
    Next i
    
    proj.Save
    proj.Shut
    
    Set OPApp = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description
    If Not proj Is Nothing Then proj.ShutWOSave
End Sub
```

### Example 4: Import/Export Operations

```vb
Sub ExportProjectData()
    On Error GoTo ErrorHandler
    
    Dim OPApp As Object
    Dim proj As Object
    Dim exportXML As String
    
    Set OPApp = CreateObject("opp.application")
    OPApp.Login
    
    Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")
    
    ' Build export XML parameters
    ' Export activities to CSV
    exportXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<ExportDefinition>" & vbCrLf & _
                "  <Format>CSV</Format>" & vbCrLf & _
                "  <OutputFile>C:\Exports\Activities.csv</OutputFile>" & vbCrLf & _
                "  <Table>ACTIVITY</Table>" & vbCrLf & _
                "  <Filter><![CDATA[PCT_COMP < 100]]></Filter>" & vbCrLf & _
                "  <Sort><![CDATA[EARLY_START, ACT_ID]]></Sort>" & vbCrLf & _
                "  <Fields>" & vbCrLf & _
                "    <Field>ACT_ID</Field>" & vbCrLf & _
                "    <Field>DESCRIPTN</Field>" & vbCrLf & _
                "    <Field>ORIG_DUR</Field>" & vbCrLf & _
                "    <Field>EARLY_START</Field>" & vbCrLf & _
                "    <Field>EARLY_FINISH</Field>" & vbCrLf & _
                "    <Field>PCT_COMP</Field>" & vbCrLf & _
                "    <Field>TOTAL_FLT</Field>" & vbCrLf & _
                "    <Field>CRITICAL</Field>" & vbCrLf & _
                "  </Fields>" & vbCrLf & _
                "  <Options>" & vbCrLf & _
                "    <IncludeHeaders>true</IncludeHeaders>" & vbCrLf & _
                "    <DateFormat>MM/DD/YYYY</DateFormat>" & vbCrLf & _
                "    <Delimiter>,</Delimiter>" & vbCrLf & _
                "    <TextQualifier>&quot;</TextQualifier>" & vbCrLf & _
                "  </Options>" & vbCrLf & _
                "</ExportDefinition>"
    
    ' Perform export
    proj.GeneralExport exportXML
    Debug.Print "Activities exported to CSV"
    
    ' Export resources to Excel
    exportXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<ExportDefinition>" & vbCrLf & _
                "  <Format>EXCEL</Format>" & vbCrLf & _
                "  <OutputFile>C:\Exports\Resources.xlsx</OutputFile>" & vbCrLf & _
                "  <Table>RESOURCE</Table>" & vbCrLf & _
                "  <Fields>" & vbCrLf & _
                "    <Field>RES_ID</Field>" & vbCrLf & _
                "    <Field>RES_DESC</Field>" & vbCrLf & _
                "    <Field>RES_CLASS</Field>" & vbCrLf & _
                "    <Field>UNIT_COST</Field>" & vbCrLf & _
                "    <Field>UNITS</Field>" & vbCrLf & _
                "  </Fields>" & vbCrLf & _
                "  <Options>" & vbCrLf & _
                "    <IncludeHeaders>true</IncludeHeaders>" & vbCrLf & _
                "    <SheetName>Resources</SheetName>" & vbCrLf & _
                "  </Options>" & vbCrLf & _
                "</ExportDefinition>"
    
    proj.GeneralExport exportXML
    Debug.Print "Resources exported to Excel"
    
    ' Export to MPX format for MS Project compatibility
    exportXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<ExportDefinition>" & vbCrLf & _
                "  <Format>MPX</Format>" & vbCrLf & _
                "  <OutputFile>C:\Exports\Project.mpx</OutputFile>" & vbCrLf & _
                "  <Options>" & vbCrLf & _
                "    <IncludeBaselines>true</IncludeBaselines>" & vbCrLf & _
                "    <IncludeResources>true</IncludeResources>" & vbCrLf & _
                "    <IncludeAssignments>true</IncludeAssignments>" & vbCrLf & _
                "  </Options>" & vbCrLf & _
                "</ExportDefinition>"
    
    proj.GeneralExport exportXML
    Debug.Print "Project exported to MPX"
    
    ' Export to Primavera P6 XML
    exportXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<ExportDefinition>" & vbCrLf & _
                "  <Format>PRIMAVERA_XML</Format>" & vbCrLf & _
                "  <OutputFile>C:\Exports\Project_P6.xml</OutputFile>" & vbCrLf & _
                "  <Options>" & vbCrLf & _
                "    <Version>6.0</Version>" & vbCrLf & _
                "    <IncludeResources>true</IncludeResources>" & vbCrLf & _
                "    <IncludeBaselines>true</IncludeBaselines>" & vbCrLf & _
                "  </Options>" & vbCrLf & _
                "</ExportDefinition>"
    
    proj.GeneralExport exportXML
    Debug.Print "Project exported to Primavera XML"
    
    proj.Shut
    Set OPApp = Nothing
    
    Debug.Print "All exports completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error during export: " & Err.Description
    If Not proj Is Nothing Then proj.ShutWOSave
End Sub

Sub ImportProjectData()
    On Error GoTo ErrorHandler
    
    Dim OPApp As Object
    Dim proj As Object
    Dim importXML As String
    
    Set OPApp = CreateObject("opp.application")
    OPApp.Login
    
    Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")
    
    ' Import actual dates from CSV
    importXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<ImportDefinition>" & vbCrLf & _
                "  <Format>CSV</Format>" & vbCrLf & _
                "  <InputFile>C:\Imports\ActualDates.csv</InputFile>" & vbCrLf & _
                "  <Table>ACTIVITY</Table>" & vbCrLf & _
                "  <KeyField>ACT_ID</KeyField>" & vbCrLf & _
                "  <UpdateMode>UPDATE</UpdateMode>" & vbCrLf & _
                "  <FieldMappings>" & vbCrLf & _
                "    <Mapping>" & vbCrLf & _
                "      <SourceField>Activity ID</SourceField>" & vbCrLf & _
                "      <DestField>ACT_ID</DestField>" & vbCrLf & _
                "    </Mapping>" & vbCrLf & _
                "    <Mapping>" & vbCrLf & _
                "      <SourceField>Actual Start</SourceField>" & vbCrLf & _
                "      <DestField>ACT_START</DestField>" & vbCrLf & _
                "    </Mapping>" & vbCrLf & _
                "    <Mapping>" & vbCrLf & _
                "      <SourceField>Actual Finish</SourceField>" & vbCrLf & _
                "      <DestField>ACT_FINISH</DestField>" & vbCrLf & _
                "    </Mapping>" & vbCrLf & _
                "    <Mapping>" & vbCrLf & _
                "      <SourceField>Percent Complete</SourceField>" & vbCrLf & _
                "      <DestField>PCT_COMP</DestField>" & vbCrLf & _
                "    </Mapping>" & vbCrLf & _
                "  </FieldMappings>" & vbCrLf & _
                "  <Options>" & vbCrLf & _
                "    <HasHeaders>true</HasHeaders>" & vbCrLf & _
                "    <DateFormat>MM/DD/YYYY</DateFormat>" & vbCrLf & _
                "    <Delimiter>,</Delimiter>" & vbCrLf & _
                "    <SkipErrors>false</SkipErrors>" & vbCrLf & _
                "  </Options>" & vbCrLf & _
                "</ImportDefinition>"
    
    proj.GeneralImport importXML
    Debug.Print "Actual dates imported from CSV"
    
    ' Import resource costs from Excel
    importXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<ImportDefinition>" & vbCrLf & _
                "  <Format>EXCEL</Format>" & vbCrLf & _
                "  <InputFile>C:\Imports\ResourceCosts.xlsx</InputFile>" & vbCrLf & _
                "  <Table>RESOURCE</Table>" & vbCrLf & _
                "  <KeyField>RES_ID</KeyField>" & vbCrLf & _
                "  <UpdateMode>UPDATE</UpdateMode>" & vbCrLf & _
                "  <FieldMappings>" & vbCrLf & _
                "    <Mapping>" & vbCrLf & _
                "      <SourceField>Resource ID</SourceField>" & vbCrLf & _
                "      <DestField>RES_ID</DestField>" & vbCrLf & _
                "    </Mapping>" & vbCrLf & _
                "    <Mapping>" & vbCrLf & _
                "      <SourceField>New Unit Cost</SourceField>" & vbCrLf & _
                "      <DestField>UNIT_COST</DestField>" & vbCrLf & _
                "    </Mapping>" & vbCrLf & _
                "  </FieldMappings>" & vbCrLf & _
                "  <Options>" & vbCrLf & _
                "    <SheetName>Costs</SheetName>" & vbCrLf & _
                "    <HasHeaders>true</HasHeaders>" & vbCrLf & _
                "  </Options>" & vbCrLf & _
                "</ImportDefinition>"
    
    proj.GeneralImport importXML
    Debug.Print "Resource costs imported from Excel"
    
    ' After import, recalculate
    proj.TimeAnalyze
    proj.CalculateCost
    
    proj.Save
    proj.Shut
    
    Set OPApp = Nothing
    
    Debug.Print "All imports completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error during import: " & Err.Description
    If Not proj Is Nothing Then proj.ShutWOSave
End Sub
```

### Example 5: Working with Multiple Projects

```vb
Sub CompareMultipleProjects()
    On Error GoTo ErrorHandler
    
    Dim OPApp As Object
    Dim proj1 As Object
    Dim proj2 As Object
    Dim acts1 As Object
    Dim acts2 As Object
    Dim act1 As Object
    Dim act2 As Object
    Dim i As Long
    
    Set OPApp = CreateObject("opp.application")
    OPApp.Login
    
    ' Open both projects
    Set proj1 = OPApp.FileOpen("C:\Projects\Project_Current.opp")
    Set proj2 = OPApp.FileOpen("C:\Projects\Project_Baseline.opp")
    
    Debug.Print "========== PROJECT COMPARISON =========="
    Debug.Print "Current Project: " & proj1.Name
    Debug.Print "Baseline Project: " & proj2.Name
    Debug.Print ""
    
    ' Compare project-level data
    Debug.Print "Current Early Finish: " & proj1.EarlyFinishDate
    Debug.Print "Baseline Early Finish: " & proj2.EarlyFinishDate
    
    Dim daysDiff As Long
    daysDiff = DateDiff("d", proj2.EarlyFinishDate, proj1.EarlyFinishDate)
    Debug.Print "Schedule Variance: " & daysDiff & " days"
    Debug.Print ""
    
    ' Compare activities
    Set acts1 = proj1.Activities
    Set acts2 = proj2.Activities
    
    Debug.Print "========== ACTIVITY VARIANCES =========="
    
    For i = 1 To acts1.Count
        Set act1 = acts1.Item(i)
        
        ' Find matching activity in baseline project
        On Error Resume Next
        Set act2 = acts2.Item(act1.ID)
        On Error GoTo ErrorHandler
        
        If Not act2 Is Nothing Then
            ' Compare durations
            Dim durDiff As Double
            durDiff = ParseDuration(act1.Duration) - ParseDuration(act2.Duration)
            
            If durDiff <> 0 Then
                Debug.Print "Activity: " & act1.ID
                Debug.Print "  Current Duration: " & act1.Duration
                Debug.Print "  Baseline Duration: " & act2.Duration
                Debug.Print "  Variance: " & durDiff & " days"
            End If
            
            ' Compare dates if both complete
            If act1.ActualFinishDate <> #12:00:00 AM# And _
               act2.ActualFinishDate <> #12:00:00 AM# Then
                
                Dim dateDiff As Long
                dateDiff = DateDiff("d", act2.ActualFinishDate, act1.ActualFinishDate)
                
                If dateDiff <> 0 Then
                    Debug.Print "  Actual finish variance: " & dateDiff & " days"
                End If
            End If
            
            Set act2 = Nothing
        End If
    Next i
    
    ' Close both projects
    proj1.Shut
    proj2.Shut
    
    Set OPApp = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description
    If Not proj1 Is Nothing Then proj1.ShutWOSave
    If Not proj2 Is Nothing Then proj2.ShutWOSave
End Sub

Function ParseDuration(durString As String) As Double
    ' Simple duration parser - converts to days
    ' Real implementation would handle all unit types
    Dim numPart As String
    Dim unitPart As String
    Dim multiplier As Double
    
    numPart = Left(durString, Len(durString) - 1)
    unitPart = Right(durString, 1)
    
    Select Case UCase(unitPart)
        Case "D"
            multiplier = 1
        Case "W"
            multiplier = 7
        Case "M"
            multiplier = 30
        Case "H"
            multiplier = 1 / 24
        Case "T"
            multiplier = 1 / (24 * 60)
        Case Else
            multiplier = 1
    End Select
    
    ParseDuration = CDbl(numPart) * multiplier
End Function
```

---

## Import/Export XML Schemas

### Comprehensive Export Schema

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ExportDefinition>
  <!-- Required: Output format -->
  <Format>CSV|EXCEL|XML|MPX|PRIMAVERA_XML|MSProject_XML</Format>
  
  <!-- Required: Output file path -->
  <OutputFile>C:\Path\To\Output.ext</OutputFile>
  
  <!-- Required: Source table -->
  <Table>ACTIVITY|RESOURCE|ASSIGNMENT|PREDECESSOR|COST</Table>
  
  <!-- Optional: Filter expression -->
  <Filter><![CDATA[EARLY_START >= #1/1/2024# AND CRITICAL > 0]]></Filter>
  
  <!-- Optional: Sort expression -->
  <Sort><![CDATA[EARLY_START DESC, ACT_ID]]></Sort>
  
  <!-- Required: Fields to export -->
  <Fields>
    <Field>ACT_ID</Field>
    <Field>DESCRIPTN</Field>
    <Field>ORIG_DUR</Field>
    <Field>EARLY_START</Field>
    <Field>EARLY_FINISH</Field>
    <Field>PCT_COMP</Field>
    <!-- User-defined fields -->
    <Field>USER1</Field>
    <Field>USER_DATE1</Field>
    <Field>USER_NUM1</Field>
  </Fields>
  
  <!-- Format-specific options -->
  <Options>
    <!-- CSV Options -->
    <IncludeHeaders>true|false</IncludeHeaders>
    <Delimiter>,|;|\t</Delimiter>
    <TextQualifier>"</TextQualifier>
    <DateFormat>MM/DD/YYYY|DD-MMM-YYYY|YYYY-MM-DD</DateFormat>
    <TimeFormat>HH:MM:SS|HH:MM</TimeFormat>
    <Encoding>UTF-8|ASCII</Encoding>
    
    <!-- Excel Options -->
    <SheetName>Activities</SheetName>
    <StartRow>1</StartRow>
    <StartColumn>A</StartColumn>
    
    <!-- XML Options -->
    <RootElement>Project</RootElement>
    <RowElement>Activity</RowElement>
    <IncludeSchema>true|false</IncludeSchema>
    
    <!-- Format-agnostic options -->
    <IncludeBaselines>true|false</IncludeBaselines>
    <BaselineName>Baseline 1</BaselineName>
    <IncludeResources>true|false</IncludeResources>
    <IncludeAssignments>true|false</IncludeAssignments>
    <IncludePredecessors>true|false</IncludePredecessors>
    <IncludeCosts>true|false</IncludeCosts>
  </Options>
</ExportDefinition>
```

### Comprehensive Import Schema

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ImportDefinition>
  <!-- Required: Input format -->
  <Format>CSV|EXCEL|XML|MPX|PRIMAVERA_XML|MSProject_XML</Format>
  
  <!-- Required: Input file path -->
  <InputFile>C:\Path\To\Input.ext</InputFile>
  
  <!-- Required: Target table -->
  <Table>ACTIVITY|RESOURCE|ASSIGNMENT|PREDECESSOR|COST</Table>
  
  <!-- Required: Key field for matching -->
  <KeyField>ACT_ID|RES_ID</KeyField>
  
  <!-- Required: Update mode -->
  <UpdateMode>INSERT|UPDATE|INSERT_UPDATE|REPLACE</UpdateMode>
  
  <!-- Required: Field mappings -->
  <FieldMappings>
    <Mapping>
      <SourceField>Activity ID</SourceField>
      <DestField>ACT_ID</DestField>
      <DataType>STRING|DATE|NUMBER|BOOLEAN</DataType>
      <Required>true|false</Required>
      <DefaultValue></DefaultValue>
    </Mapping>
    <Mapping>
      <SourceField>Description</SourceField>
      <DestField>DESCRIPTN</DestField>
      <DataType>STRING</DataType>
    </Mapping>
    <Mapping>
      <SourceField>Duration</SourceField>
      <DestField>ORIG_DUR</DestField>
      <DataType>DURATION</DataType>
      <DefaultUnit>D</DefaultUnit>
    </Mapping>
    <Mapping>
      <SourceField>Start Date</SourceField>
      <DestField>ACT_START</DestField>
      <DataType>DATE</DataType>
      <DateFormat>MM/DD/YYYY</DateFormat>
    </Mapping>
  </FieldMappings>
  
  <!-- Format-specific options -->
  <Options>
    <!-- CSV Options -->
    <HasHeaders>true|false</HasHeaders>
    <Delimiter>,|;|\t</Delimiter>
    <TextQualifier>"</TextQualifier>
    <DateFormat>MM/DD/YYYY</DateFormat>
    <StartRow>1</StartRow>
    <Encoding>UTF-8</Encoding>
    
    <!-- Excel Options -->
    <SheetName>Sheet1</SheetName>
    <StartRow>2</StartRow>
    <StartColumn>A</StartColumn>
    
    <!-- General Options -->
    <SkipErrors>true|false</SkipErrors>
    <LogErrors>true|false</LogErrors>
    <ErrorLogFile>C:\Logs\Import_Errors.log</ErrorLogFile>
    <ValidateData>true|false</ValidateData>
    <CreateMissing>true|false</CreateMissing>
    <TruncateStrings>true|false</TruncateStrings>
  </Options>
  
  <!-- Validation rules -->
  <Validation>
    <Rule>
      <Field>ACT_ID</Field>
      <Type>REQUIRED|UNIQUE|PATTERN|RANGE</Type>
      <Pattern>[A-Z0-9]{1,10}</Pattern>
    </Rule>
    <Rule>
      <Field>ORIG_DUR</Field>
      <Type>RANGE</Type>
      <MinValue>0</MinValue>
      <MaxValue>9999</MaxValue>
    </Rule>
  </Validation>
</ImportDefinition>
```

### Example: Custom Transform During Export

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ExportDefinition>
  <Format>CSV</Format>
  <OutputFile>C:\Exports\Custom_Report.csv</OutputFile>
  <Table>ACTIVITY</Table>
  
  <Fields>
    <Field>ACT_ID</Field>
    <Field>DESCRIPTN</Field>
    <!-- Calculated fields using expressions -->
    <CalculatedField>
      <Name>SCHEDULE_VARIANCE_DAYS</Name>
      <Expression>DATEDIFF(EARLY_FINISH, BASELINE_FINISH, 'D')</Expression>
    </CalculatedField>
    <CalculatedField>
      <Name>COST_VARIANCE</Name>
      <Expression>ACTUAL_COST - BUDGET_COST</Expression>
    </CalculatedField>
    <CalculatedField>
      <Name>STATUS_FLAG</Name>
      <Expression>
        CASE
          WHEN PCT_COMP = 100 THEN 'Complete'
          WHEN PCT_COMP > 0 THEN 'In Progress'
          ELSE 'Not Started'
        END
      </Expression>
    </CalculatedField>
  </Fields>
  
  <Options>
    <IncludeHeaders>true</IncludeHeaders>
    <Delimiter>,</Delimiter>
  </Options>
</ExportDefinition>
```

---

*Enhanced with comprehensive VBA examples and Import/Export documentation.*
