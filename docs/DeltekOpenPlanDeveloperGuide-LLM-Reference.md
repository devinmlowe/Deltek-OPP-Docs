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

## Custom Import/Export Scripts (Transfer.dat)

### Overview

When you use General Import/Export from the Integration tab, Open Plan uses scripts defined in the **Transfer.dat** file (located in the local system folder). This is a text file that can be edited to create custom import/export specifications.

**CRITICAL:** All import/export command scripts are **case-sensitive**. Commands must be entered in UPPERCASE or you may receive errors.

### Transfer.dat File Structure

Transfer.dat can contain:
- Script definitions for import/export operations
- Pointers to separate include files
- Each line must start with a command word

### Default Import Scripts

**Activity Information** - Creates activities with:
- ID
- Description
- Early Dates, Late Dates, Scheduled Dates
- User Character Fields 1 and 2

**Import Resource Actual Time and Cost** - Imports to CST table:
- ACT_ID, RES_ID
- ACWP_QTY (Actual Quantity of Work Performed)
- ACWP_CST (Actual Cost of Work Performed)
- START_DATE, END_DATE (Period dates)
- NOTE: Adds new records, does not update existing

**Simple Activity Table Import** - Comma-delimited import for:
- OPP_ID, DESCRIPTION, ORIG_DUR
- ESDATE, EFDATE, COMPSTAT
- NOTE: Use "Add data to open project" option to append without overwriting

### Default Export Scripts

**Baseline Export to Cobra (Activities)** - Transaction file for activity baseline data to Cobra

**Baseline Export to Cobra (Assignments)** - Transaction file for resource assignment baseline data to Cobra

**Cobra Status Update** - Transaction file for status update to Cobra

**Date & Status Report (XML)** - Activity date/status to XML with Actdata2.xsl stylesheet

**Export Resource Actual Time and Cost** - CSV file from CST table fields:
- ACT_ID, RES_ID, ACWP_QTY, ACWP_CST, START_DATE, END_DATE

**GRANEDA** - Export in American Netronic's GRANEDA format

**IPMR Format 6 Report (UN/CEFACT)** - US Government IPMR Format 6 requirements

**IPMR Format 6 Report with Risk (UN/CEFACT)** - IPMR Format 6 with optional risk data (optimistic, pessimistic, most likely durations)

**Predecessor & Successor Report (XML)** - Activity and relationship data using Actdata1.xsl

**Resource/Activity Report (XML)** - Activities and resource assignments using Rdsdata.xsl

**Simple Activity Table Export** - CSV from OPP_ACT table fields:
- OPP_ID, DESCRIPTION, ORIG_DUR, ESDATE, EFDATE, COMPSTAT

**Primavera XER Format Standard** - P6 .xer file with Code fields 1-20

**Primavera XER with WBS in C90** - P6 .xer file with Code fields 1-20 and Code field 90

### Basic Import/Export Script Example

```
EXPORT csv Simple Activity Export to Excel
REM exports ID,Desc,Duration,ESDate,EFDate,Activity Cost
TABLE ACT
SORT ACT_ID
FIELD ACT_ID
FIELD DESCRIPTION
FIELD ORIG_DUR
FIELD ESDATE

IMPORT csv Simple Activity Import from Excel
REM imports ID, Description, and Duration
TABLE ACT
FIELD ACT_ID
FIELD DESCRIPTION
FIELD ORIG_DUR

EXPORT xml Example XML Activity Export 1
INCLUDE ACTDATA1.XFR
```

### Core Commands

#### IMPORT Command

Syntax: `IMPORT <extension> <name>`

Defines start of import script. Everything from this command to next IMPORT/EXPORT or EOF defines the script.

- `<extension>`: Default file extension (can use wildcard `*`)
- `<name>`: Appears in Open Plan Import dialog

Examples:
```
IMPORT csv Excel file
IMPORT * Excel
```

#### EXPORT Command

Syntax: `EXPORT <extension> <name>`

Defines start of export script. Everything from this command to next IMPORT/EXPORT or EOF defines the script.

- `<extension>`: Default file extension (can use wildcard `*`)
- `<name>`: Appears in Open Plan Export dialog

Examples:
```
EXPORT XML Extended Markup Language (HTML)
EXPORT * XML
```

#### TABLE Command

Syntax: `TABLE <optabletype>`

Indicates the current table for the script. Everything from this to next TABLE, IMPORT, EXPORT, or EOF applies to this table.

Table types: ACT, REL, ASG, CST, USE, RES, AVL, CLH, CLR, SCA, CNN, PRJ, etc.

Example:
```
TABLE ACT
```

#### RECORD_TYPE Command

Syntax: `RECORD_TYPE <recordtype>`

Defines identifier for record type, assumed to be first item of each record.

#### FIELD Command

Syntax: `FIELD <opfieldname> [<width>|<header>][<format>]`

Designates next field from current table to be processed.

- `<width>`: For fixed format, field width
- `<header>`: For delimited with header, field identifier
- `<format>`: Optional format for numeric/date/enumerated data

**Numeric Format Conventions:**

| Format | Result |
|--------|--------|
| ( | If negative, suppress minus sign and print left parenthesis |
| ) | If negative, print right parenthesis |
| 9 | Print digit, replace leading zeros with spaces |
| 0 | Print digit, replace leading spaces with zeros |
| _ | Print digit, suppress leading spaces |
| . or , | Insert period or comma |
| %X | Insert local currency symbol |
| %% | Insert percent sign |
| %( | Insert left parenthesis |
| %) | Insert right parenthesis |

**Format Examples:**

Assume field PPC contains value 27.42:

| Command | Result |
|---------|--------|
| FIELD PPC 10 __________ | 27^^^^^^^ |
| FIELD PPC 10 0000000000 | 0000000027 |
| FIELD PPC 10 9999999999 | ^^^^^^^^27 |
| FIELD PPC 10 999,999.99 | ^^^^^27.42 |
| FIELD PPC 10 %X________ | $27^^^^^^^ |

**Date Format:** Use `%M%D%Y` produces dates like 123103

**Enumerated Field:** Use `S` for short form: `FIELD REL_TYPE S` exports "FS" instead of "Finish To Start"

#### INCLUDE Command

Syntax: `INCLUDE <filename>`

Allows script files to be included by reference. Include files may be nested but must NOT contain IMPORT or EXPORT commands.

Example:
```
INCLUDE ACTDATA1.XFR
```

### Additional Commands

#### ADD_MISSING_KEYS

For imports, automatically create records for missing data. Example: If import has resource assignments but resource file is empty, Open Plan creates the referenced resource records automatically.

```
ADD_MISSING_KEYS
```

#### DATE_FORMAT

Syntax: `DATE_FORMAT <format>`

Defines date format using Open Plan date format strings.

Examples:
```
DATE_FORMAT %M/%D/%Y
DATE_FORMAT "%M/%D/%C %H:%T"  'With space requires quotes
```

Results:
| Command | Result |
|---------|--------|
| DATE_FORMAT %M/%D/%C_%H:%T | 06/02/2003_08:00 |
| DATE_FORMAT %M/%D/%C %H:%T | 06/02/2003 |
| DATE_FORMAT "%M/%D/%C %H:%T" | 06/02/2003 08:00 |

**IMPORTANT:** Open Plan requires **two-digit** month, day, and time formats when importing from comma-delimited files. Example: Use `02/05/04` not `2/5/04`.

#### DELIMITED

Syntax: `DELIMITED <character>`

Specifies delimiter character. Default is comma if not specified.

Special characters:
- `\t` - Tab delimiter
- `\0` - Binary zero delimiter

Examples:
```
DELIMITED ,
DELIMITED ;
DELIMITED \t
```

#### DURATION_FORMAT

Syntax: `DURATION_FORMAT <durunit> [<durstring>]`

Defines duration format for export.

Duration units:
- m = months
- w = weeks
- d = days
- h = hours
- t = minutes

Examples:
```
DURATION_FORMAT h              'Round to hours
DURATION_FORMAT thdwm          'Use custom abbreviations
DURATION_FORMAT *thdwm         'Custom units, no rounding
DURATION_FORMAT !              'Export zeros with default units
DURATION_FORMAT !thdwm         'Export zeros with custom units
```

#### FIELD_SPECIAL

Syntax: `FIELD_SPECIAL <specialfieldname>`

Transfers information not stored in a specific field.

**For any table:**
- $ROW_NUMBER$ (export only)
- $PROJECT_NAME$

**For activity table:**
- $MPX_PREDS$
- $MPX_PREDS_REPLACES$
- $MPX_PRED_ROWS$
- $MPX_TARGET_DATE$
- $MPX_TARGET_TYPE$
- $MPX_FIXED$
- $BAAN_CODE_NUMBER$
- $BAAN_CODE_VALUE$
- $BAAN_CODE_DESC$

**For code tables:**
- $CODE_NUMBER$
- $BAAN_HIERARCHY_PARENT$
- $BAAN_CODE_NUMBER$
- $BAAN_CODE_NAME$
- $BAAN_CODE_VALUE$
- $BAAN_CODE_DESC$

**For resource description table:**
- $BAAN_HIERARCHY_PARENT$
- $MPX_RES_UNIQUEID$
- $MPX_RES_LEVEL$
- $MPX_CAL$

Special fields output constants: Anything not matching special field pattern is treated as literal. To output literal "$", use "$$".

Example:
```
FIELD_SPECIAL $PROJECT_NAME$
FIELD_SPECIAL $ROW_NUMBER$
```

#### FILTER

Syntax: `FILTER [<filtername>] [<filterdefinition>]`

Applies filter to export data. Filter definition should contain no spaces.

Example:
```
FILTER CRITICAL>0ANDPCT_COMP<100
```

#### FIXED

Syntax:
```
FIXED <fieldwidth>
FIXED <margin> <fieldwidth>
```

Defines file as fixed-format.

- One parameter: `<fieldwidth>` = characters for data record type
- Two parameters: `<margin>` = characters to ignore before record type, `<fieldwidth>` = record type width

Must precede first TABLE command. If absent, file assumed comma-delimited.

#### HEADER

Syntax: `HEADER <headerrecordtype>`

Defines record type of header record. Header records indicate fields to import/export and their order using identifiers matching `<header>` parameter of FIELD command.

#### LINK

Syntax: `LINK [<fieldname>]`

Links current table to previously defined table.

Example - link assignment table to activity table:
```
LINK ID
```

Special form without field name links to most recently defined table of same type on record-by-record basis.

Tables may be linked to linked tables, creating chains of any length.

#### LITERAL_END

Syntax: `LITERAL_END`

Terminates literal data introduced by LITERAL_HEADER or LITERAL_FOOTER.

#### LITERAL_FOOTER

Syntax: `LITERAL_FOOTER`

Introduces lines to transfer at end of output for current table. Terminated by LITERAL_END.

Example:
```
LITERAL_FOOTER
</ACTIVITIES>
</PROJECT>
LITERAL_END
```

#### LITERAL_HEADER

Syntax: `LITERAL_HEADER`

Introduces lines to transfer at beginning of output for current table. Terminated by LITERAL_END.

Example:
```
LITERAL_HEADER
<?xml version="1.0" ?>
<PROJECT>
<ACTIVITIES>
LITERAL_END
```

**NOTE:** LITERAL commands apply only to exports. Can be defined outside table definition scope.

#### MPX_CALENDAR_DEFINITION

Syntax: `MPX_CALENDAR_DEFINITION`

Processes calendar definitions in format required by .mpx record type 20. Used only with CLH table type.

#### MPX_CALENDAR_HOURS

Syntax: `MPX_CALENDAR_HOURS`

Processes standard working hours in format required by .mpx record type 25. Used only with CLH table type.

#### MPX_CALENDAR_EXCEPTIONS

Syntax: `MPX_CALENDAR_EXCEPTIONS`

Processes calendar exceptions (holidays) in format required by .mpx record type 26. Used only with CLH table type.

#### REM

Syntax: `REM <comment>`

Adds comment lines to script. Open Plan ignores lines starting with REM.

Example:
```
REM This exports critical activities only
```

#### SKIP

Syntax: `SKIP <numberoffields>|<numberofcharacters>`

Skips items before next defined field.

- Delimited: `<numberoffields>` to skip
- Fixed-format: `<numberofcharacters>` to skip

Typically used to skip unnecessary fields during import.

#### SORT

Syntax: `SORT [-]<sort_field1>[,[-]<sort_field2>[,...]]`

Applies sort to export data. Negative sign indicates descending sort.

Example:
```
SORT ACT_ID
SORT C1.codedesc,(ssdate-esdate)
```

#### UPDATE

Syntax: `UPDATE`

For import, specifies existing records matching incoming ones should be updated. Incoming data not matching existing records are added as new.

Apply to each table that needs updating.

#### UPDATE_ONLY

Syntax: `UPDATE_ONLY`

For import, incoming data only updates existing records. Incoming data without match is ignored.

**IMPORTANT:** When tables are parent-child linked (e.g., activities and assignments):
- If parent is in UPDATE or UPDATE_ONLY mode and child is not, all existing children are deleted before adding new child records
- When LINK used on same table type, second instance automatically set to UPDATE_ONLY mode

#### UPDATE_REMAINING

Syntax: `UPDATE_REMAINING`

For CST table imports only. Updates Remaining Quantity on activity Resource Assignments when Actual Quantities imported for Activity Resources where resource configured for manual progress (Progress Based On Activity Progress disabled).

### Table Types

3-character table identifiers corresponding to database extensions:

| Table Type | Description |
|------------|-------------|
| ACT | Activity details |
| REL | Activity relationships |
| ANN | Baseline activities (A01 for baseline 1) |
| UNN | Baseline resource usage (U01 for baseline 1) |
| ASG | Resource assignments |
| CST | Resource actuals |
| USE | Resource usage |
| RSK | Risk key activities |
| RES | Resource definitions |
| AVL | Resource availabilities |
| RSL | Resource cost escalation (ESC synonym) |
| CLH | Calendar header (CAL synonym) |
| CLR | Calendar detail |
| SCA | Code structure association (CDH synonym) |
| CNN | Any number of attached code tables |
| <nn> | Specific code table number |
| PRJ | Project directory details |

**PRJ Table Field Name Mappings (backward compatibility):**
- PROJECT_NAME → DIR_ID
- PATH_NAME → DIR_ID
- CALENDAR_PATH_NAME → CLD_ID
- RESOURCE_PATH_NAME → RDS_ID
- DESCRIP → DESCRIPTION
- MANAGER → OPMANAGER
- COMPANY → OPCOMPANY
- CLIENT → OPCLIENT

Access PROJ_NOFCODES using:
```
FIELD_SPECIAL $PRJ_NOFCODES$
```

**CDH Pseudo-Table Fields:**
- CODE_NUMBER
- CODE_NAME → COD_ID
- PATH_NAME → COD_ID
- DESCRIP → DESCRIPTION

**CNN Table:** Imports/exports any number of code files with same record type. Use `$CODE_NUMBER$` to distinguish different code files, or use specific `<nn>` table type for specific code files.

**NOTE:** Cannot use both CODE_NUMBER and CODE_NAME to import code information.

### XML Import/Export Commands

#### ATTRIBUTE Keyword

Synonym for FIELD, supports XML nomenclature to distinguish from ELEMENT.

#### ELEMENT Keyword

Similar to FIELD but causes field to be transferred as XML element rather than attribute.

Example output:
```xml
<ACTIVITY <ID>1</ID>
DESCRIPTION="Environment Management System"
ESDATE="02Jan97" />
```

**NOTE:** Elements always output first regardless of definition order.

#### HIERARCHICAL Keyword

Causes export to be nested using hierarchical key. Works for activity, resource description, and code tables. Linked tables also nested.

Example script:
```
EXPORT XML XML Export (With Style Sheet)
DATE_FORMAT %D%A%Y
LITERAL_HEADER
<?xml version="1.0" ?>
LITERAL_END
STYLESHEET op1.xsl

TABLE ACT
HIERARCHICAL
XML_TAG ACTIVITIES
XML_TAG ACTIVITY
FIELD ID ID
FIELD ACT_DESC DESCRIPTION
FIELD ESDATE ESDATE

TABLE REL
XML_TAG PREDECESSORS
XML_TAG PREDECESSOR
LINK SUCC_ID
FIELD PRED_ACT_UID PRED_ACT_UID

TABLE REL
XML_TAG SUCCESSORS
XML_TAG SUCCESSOR
LINK PRED_ACT_UID
FIELD SUCC_ID SUCC_ID
```

**IMPORTANT:** On import, nesting does not imply hierarchy - hierarchy must be explicit in activity ID.

#### STYLESHEET Keyword

Defines stylesheet to format XML data in browser. Appears after LITERAL_END following XML definition.

Text following STYLESHEET should be filename or URL:
- Local file (no slashes): Interpreted as in Open Plan global directory
- Current directory: Use ".\filename"
- Absolute path: Use full path or URL

Example:
```
STYLESHEET op1.xsl
```

No effect on import.

#### XML_TAG Keyword

Defines XML import/export. Can appear up to twice per table:
- First: Tag for collection
- Second: Tag for each element in collection

If only one present, applies to each element.

Cannot use with: RECORD_TYPE, FIXED, HEADER, MPX_CALENDAR_* commands

All XML documents must have root element.

Example:
```
TABLE ACT
XML_TAG ACTIVITIES
XML_TAG ACTIVITY
```

#### FIELD Keyword (XML Context)

If XML_TAG specified for table, second parameter on FIELD contains XML attribute name. May be omitted if same as Open Plan field name.

**Export:** Blank character string attributes not exported

**Import:**
- Not all attributes required for every element
- At least one attribute needed or no record created
- Even ID can be missing (uses auto-numbering)
- Attributes in input file but not defined in script are ignored (no log message)
- Relationships cannot be created prior to both activities

#### LINK Keyword (XML Context)

Embeds table inside parent table element.

Special form:
```
LINK PRJ
```

Embeds elements inside project element (root element).

**NOTE:** In valid XML, all data must appear within root tag. Every table except PRJ must contain LINK keyword linking to PRJ or previously-defined table.

#### LITERAL_HEADER & LITERAL_FOOTER (XML Context)

Can be defined outside table definition scope. Facilitates XML declaration and heading information.

### Import/Export Considerations

1. Log messages may reference lines that don't uniquely identify issue
2. On input, can nest top-level tags within other tags (works but doesn't imply hierarchy)
3. On export, no way to nest activities within other activities
4. Import log messages may not pinpoint error location (CR/LF has no significance in XML)

### XML Example 1 - Complete Script

```
EXPORT XML XML Export
LITERAL_HEADER
<?xml version="1.0" ?>
LITERAL_END

TABLE PRJ
XML_TAG PROJECT

TABLE ACT
XML_TAG ACTIVITIES
XML_TAG ACTIVITY
LINK PRJ
FIELD ID ID
FIELD ACT_DESC DESCRIPTION

TABLE REL
XML_TAG PREDECESSORS
XML_TAG PREDECESSOR
LINK SUCC_ID
FIELD PRED_ACT_UID PRED_ACT_UID

TABLE REL
XML_TAG SUCCESSORS
XML_TAG SUCCESSOR
LINK PRED_ACT_UID
FIELD SUCC_ID SUCC_ID
```

**Output Example:**
```xml
<?xml version="1.0" ?>
<PROJECT>
<ACTIVITIES>

<ACTIVITY
   ID="1.01.02"
   DESCRIPTION="Site Management" >

   <SUCCESSORS>
            <SUCCESSOR
                 SUCC_ID="1.01.01" />
   </SUCCESSORS>

</ACTIVITY>

<ACTIVITY
   ID="1.02"
   DESCRIPTION="Coordination Planning" >

   <PREDECESSORS>
          <PREDECESSOR
                 PRED_ACT_UID="1.01" />
   </PREDECESSORS>

   <SUCCESSORS>
           <SUCCESSOR
                SUCC_ID="1.03" />
   </SUCCESSORS>

</ACTIVITY>

</ACTIVITIES>
</PROJECT>
```

**Notes:**
1. XML declaration line introduces file
2. PROJECT defines root element (required for multiple tables)
3. Exporting both successors and predecessors ensures all relationships imported regardless of activity order
4. Clicking XML file invokes web browser with hierarchical display

### XML Example 2 - Relationships as Separate Collection

```
LITERAL_HEADER
<?xml version="1.0" ?>
LITERAL_END

TABLE PRJ
XML_TAG PROJECT

TABLE ACT
XML_TAG ACTIVITIES
XML_TAG ACTIVITY
LINK PRJ
FIELD ID ID
FIELD ACT_DESC DESCRIPTION

TABLE REL
XML_TAG PREDECESSORS
XML_TAG PREDECESSOR
LINK PRJ
FIELD SUCC_ID SUCC_ID
FIELD PRED_ACT_UID PRED_ACT_UID
```

**Output Example:**
```xml
<?xml version="1.0" ?>
<PROJECT>

<ACTIVITIES>

   <ACTIVITY
      ID="1"
      DESCRIPTION="Environmental Management System" />

   <ACTIVITY
      ID="1.01"
      DESCRIPTION="Requirements Development"/>

   <ACTIVITY
      ID="1.01.01"
      DESCRIPTION="Req. Coordination &amp; Planning" />

</ACTIVITIES>

<PREDECESSORS>

   <PREDECESSOR
      SUCC_ID="1.01.01"
      PRED_ACT_UID="1.01.02" />

   <PREDECESSOR
      SUCC_ID="1.02"
      PRED_ACT_UID="1.01" />

   <PREDECESSOR
      SUCC_ID="1.02.02"
      PRED_ACT_UID="1.02.01" />

</PREDECESSORS>
</PROJECT>
```

**Notes:**
1. Project tag required as root (contains both activities and relationships)
2. `&amp;` represents ampersand in XML (Open Plan translates automatically)

**XML Special Characters:**
- `&amp;` for &
- `&lt;` for <
- `&gt;` for >
- `&quote;` for "
- `&apos;` for '

### Important Import/Export Notes

**Problem Characters:** These symbols may cause issues as delimiters in fields:
- Quotation marks (")
- Commas (,)
- Piping symbols (|)
- Semicolons (;)

**Syntax Conventions:**
- `<...>` Required parameters
- `[...]` Optional parameters
- `[...|...]` Choice of optional parameters

**Best Practices:**
1. Always use UPPERCASE for commands
2. Test scripts on sample data first
3. Use REM for documentation
4. Include header records for clarity
5. Validate date formats (two-digit month/day required)
6. Check for special characters in data
7. Use filters to limit scope
8. Log errors for troubleshooting
9. Create backups before imports

---

## Calculated Fields

### Overview

User-defined calculated fields allow you to calculate and display data not stored in standard project database tables. With calculated fields, you can:
- Extend flexibility of any view by displaying custom calculations
- Display calculated columns in spreadsheet views
- Show calculations in activity boxes
- Include in custom filter and sort expressions
- Use in global edit expressions

**NOTE:** Macros `<USERDIR>` and `<SYSDIR>` may be used to shorten path names for executables or documents.

**Accessibility:**
- **Shared items** (all users): Store AddIns.dat in Open Plan executable folder
- **User-specific items** (single user): Store in user's Open Plan User folder
- User-specific tools appear after shared tools in lists

### Elements of Calculated Field Expressions

Expressions can include:
- Constants (text, date, numeric, logical, enumerated)
- Field names (from current or linked tables)
- Functions (extensive library detailed below)
- Other calculated fields
- Mathematical operators (+, -, *, /, ^, **)
- Character operators (+, $)
- Duration operators (+, -, *, /)
- Relational operators (=, <>, >, >=, <, <=, $)
- Logical operators (AND, OR, NOT, AND NOT)
- User-Defined Variables (BEGIN VARIABLES/END VARIABLES blocks)

### Constants

| Constant Type | Syntax | Examples |
|---------------|--------|----------|
| Text | Enclose in single or double quotes | "Programmers", 'Phase I' |
| Dates | Enclose in curly brackets { } | {01JAN01}, {12/01/01} |
| Durations | Enclose in pipe characters \| \| | \|4h\|, \|3.5d\| |
| Numerics | Positive or negative, with/without decimals | 1, 320000, 12.1, -123.78 |
| Logical | Boolean values | [TRUE], [FALSE] |
| Enumerated | Enclose in square brackets [ ] | [ASAP], [ALAP], [Start Milestone], [Finish Milestone], [Discontinuous], [Subproject], [Hammock], [Effort Driven], [External Subproject] |

**Enumeration Types:**
- ACTS — Activity Status
- ACTT — Activity Type
- BOOL — Boolean
- CRIT — Critical Flag
- CURV — Curve
- DIST — Risk Distribution
- EVTE — Earned Value Technique
- LOGI — Activity Logic Flag
- PRJS — Project Status
- PROG — Progress
- RELT — Relationship type
- RESC — Resource class
- RSCL — Resource scheduling type
- TARG — Target type

### Field Names

**IMPORTANT:** Use field names, NOT descriptive column headings.
- Correct: `ESDATE` (not "Early Start")
- Correct: `DESCRIPTN` (not "Description")

**Linked Tables:** Use linking field name + period + field name:
- `C1.DESCRIPTION` — Description for code in C1 field
- `RES_ID.DESCRIPTION` — Description for resource ID
- Linking fields shown with double chevron (») in Fields dialog

**WARNING:** Calculated field cannot reference itself (circular reference).

### Operators

**Mathematical:**
- Add: `+`
- Subtract: `-`
- Multiply: `*`
- Divide: `/`
- Group: `( )`
- Exponentiate: `^` or `**`

**Precedence Rules:**
1. Grouping operations first
2. Multiplication and division second
3. Addition and subtraction last

**Character:**
- Concatenate: `+`
- Is contained in: `$` (returns logical result)

**Duration Operations:**

| Operation | Result Type |
|-----------|-------------|
| Duration + Duration | Duration |
| Duration + Date | Date |
| Duration - Duration | Duration |
| Date - Duration | Date |
| Duration / Duration | Decimal |
| Duration / Number | Duration |
| Duration * Number | Duration |

**Relational:**
- Equal to: `=`
- Not equal to: `<>`
- Greater than: `>`
- Greater than or equal to: `>=`
- Less than: `<`
- Less than or equal to: `<=`
- Contained in: `$`

**Logical:**
- `AND`
- `OR`
- `NOT`
- `AND NOT`
- Grouping: `()`

### User-Defined Variables

Open Plan optimizes expressions by identifying repeated subexpressions. For explicit control and readability, use variable blocks.

**Rules:**
- Variables defined in BEGIN VARIABLES/END VARIABLES block
- Each definition on own line
- Result type determined automatically
- Variable names: no spaces, avoid field/function name conflicts
- Not case sensitive
- Format: `<variable_name> = <expression>`
- Can reference previously defined variables

**Example:**
```
BEGIN VARIABLES
X = DATEDIFFERENCE(ASDATE, TIMENOW())
Y = IIF(x<=|2d|, IIF(x<=|1d|, 2, 1), -1)
END VARIABLES

IIF(Y>0, "OK", "Warning")
```

**Benefits:**
1. More readable and maintainable
2. Faster parsing of complex expressions
3. Subexpression evaluated once per cell

### Calculated Field Functions

**Complete Function List:**

| Category | Functions |
|----------|-----------|
| **Mathematical** | ABS(), MAX(), MIN(), ROUND(), SQRT() |
| **Date/Time** | DATE(), DATEADD(), DATEDIFFERENCE(), DATEFORMAT(), DAY(), DOW(), CDOW(), MONTH(), CMONTH(), YEAR(), GO_MONTH(), TIMENOW(), FISCALPERIOD() |
| **Duration** | DURATION() |
| **String** | LEFT(), RIGHT(), MID(), SUBSTR(), LEN(), INSTR(), LTRIM(), TRIM(), UPPER(), LOWER(), SPACE(), STR(), STRTRAN(), STUFF(), NEWLINE() |
| **Conversion** | CTOD(), VAL(), NUMBER_FORMAT() |
| **Logical** | IIF(), INLIST(), HAS_NOTE(), FAIL_EVALUATE() |
| **Data Access** | GET_FIELD(), BASELINE_FIELD(), GET_FIRST_RECORD_IN_SUMMARY(), GET_NOTE(), GET_ASSGNS(), GET_PREDS(), GET_SUCCS(), GET_CHILDREN(), GET_COSTS(), GET_RISKS(), GET_USAGES(), GET_RELATED(), GET_RELATED_COUNT() |
| **Hierarchy** | LEVEL(), LOCAL(), PARENT() |
| **System** | USER_ID(), RECORD_NUMBER(), EVAL(), FORMAT_HEADING_ITEM() |

### Function Reference (Detailed)

#### ABS()

**Purpose:** Returns absolute value of numeric variable

**Data Type:** Decimal or integer

**Syntax:** `ABS(<value>)`

**Example:** `ABS(-4)` returns `4`

---

#### BASELINE_FIELD()

**Purpose:** Returns field from baseline directory table (OPP_BAS) for selected baseline

**Data Type:** Native type of field requested

**Syntax:** `BASELINE_FIELD(<SelectedBaselineIndex>, <BaselineFieldName>)`

**Parameters:**
- SelectedBaselineIndex: 1-3 (or 0 = same as 1)
- BaselineFieldName: Quote-delimited field name
  - "BASETYPE" (0=early, 1=late, 2=schedule)
  - "BAS_ID" (baseline name)
  - "DESCRIPTION"

**Example:** `BASELINE_FIELD(1, "DESCRIPTION")` returns "Performance Measurement Baseline" if baseline at index 1 is "PMB"

**Tables:** Activity, Relationship, Assignment, Cost, Risk, Subproject, Project Directory

---

#### CDOW()

**Purpose:** Returns full day of week

**Data Type:** Character

**Syntax:** `CDOW(<date>)`

**Example:** `CDOW({07OCT04})` returns "Tuesday"

---

#### CMONTH()

**Purpose:** Returns full month name

**Data Type:** Character

**Syntax:** `CMONTH(<date>)`

**Example:** `CMONTH({07OCT04})` returns "October"

---

#### CTOD()

**Purpose:** Converts character string to DATE data type

**Data Type:** Date

**Syntax:** `CTOD(<String Expression>)`

**Example:** `CTOD(STR(USER_NUM01) + "/" + STR(USER_NUM02) + "/" + STR(YEAR(TIMENOW())))`

If USER_NUM01=12, USER_NUM02=31, Time Now=1/1/2006, returns {12/31/2006}

---

#### DATE()

**Purpose:** Returns current date

**Data Type:** Date

**Syntax:** `DATE()`

**Example:** `DATE()` returns current date (e.g., 04OCT04)

---

#### DATEADD()

**Purpose:** Adds duration to date using calendar

**Data Type:** Date

**Syntax:** `DATEADD(<start date>, <duration>, <calendar>)`

**Parameters:**
- start date: Date field or constant
- duration: Duration field or constant
- calendar: Optional calendar name
  - If omitted: Uses project's <Default> calendar or 40-hour week

**Example:** `DATEADD(ESDATE, |2d|, "CAL1")` adds 2 days to ESDATE using CAL1

**Special:** Use field name `CLH_ID` for activity's own calendar (or `ID.CLH_ID` for assignment table)

---

#### DATEDIFFERENCE()

**Purpose:** Returns difference between two dates as duration

**Data Type:** Duration

**Syntax:** `DATEDIFFERENCE(<date1>, <date2>, <calendar>)`

**Parameters:**
- date1, date2: Date fields or constants
- calendar: Optional (uses project calendar if omitted)

**Example:** `DATEDIFFERENCE({07OCT04}, {08OCT04}, "CAL1")` returns `2d`

---

#### DATEFORMAT()

**Purpose:** Returns date in specified format

**Data Type:** Character

**Syntax:** `DATEFORMAT(<date field>, <format string>)`

**Example:** `DATEFORMAT(ESDATE, "%D%A%Y")` returns "07OCT04"

**NOTE:** Uses same format strings as Date Scale Preferences dialog

---

#### DAY()

**Purpose:** Returns day of month (1-31)

**Data Type:** Integer

**Syntax:** `DAY(<date>)`

**Example:** `DAY({04OCT04})` returns `4`

---

#### DOW()

**Purpose:** Returns day of week (1-7), based on .ini file setting

**Data Type:** Integer

**Syntax:** `DOW(<date>)`

**Example:** `DOW({07OCT04})` returns `3` (Tuesday)

---

#### DURATION()

**Purpose:** Returns number of minutes for given duration

**Data Type:** Integer

**Syntax:** `DURATION(<Duration Value>)`

**Example:** `DURATION(|2d|)` returns `960` (8 hours × 2 days × 60 min/hour)

---

#### EVAL()

**Purpose:** Evaluates string expression as dynamic calculated field

**Data Type:** Character

**Syntax:** `EVAL(<exp>)`

**Example:** If USER_CHR01="ESDATE", USER_CHR02="EFDATE", ESDATE={04OCT04}, EFDATE={06OCT04}:
```
EVAL('CTOD(USER_CHR02) - CTOD(USER_CHR01)')
```
Returns duration `2d`

---

#### FAIL_EVALUATE()

**Purpose:** Tests if expression would fail/return blank

**Data Type:** Boolean

**Syntax:** `FAIL_EVALUATE(<String Expression>)`

**Example:** `FAIL_EVALUATE("C25")` returns TRUE if no code file at index 25

**Note:** Different from "C25 IS_EMPTY" which also returns TRUE if C25 exists but is blank

---

#### FISCALPERIOD()

**Purpose:** Returns label for fiscal reporting period containing date

**Data Type:** Character

**Syntax:** `FISCALPERIOD(<Date>, <Reporting Calendar>)`

**Parameters:**
- Date: Search date
- Reporting Calendar: Optional (defaults to project reporting calendar)

**Example:** `FISCALPERIOD(ESDATE)` or `FISCALPERIOD(ESDATE, "REPORTING_CALENDAR_NAME")`

Returns LABEL from reporting calendar containing supplied date

---

#### FORMAT_HEADING_ITEM()

**Purpose:** Creates custom grouping heading with summary rows

**Data Type:** Character

**Syntax:** `FORMAT_HEADING_ITEM(<Expression>, <WidthInCharUnits>, <Wordwrap>, <Summarize>)`

**Parameters:**
- Expression: Character string to format
- WidthInCharUnits: Width in character units
- Wordwrap: 0 (no wrap) or 1 (wrap active)
- Summarize: Optional - 0 (don't summarize) or 1 (summarize), default=1

**Example:**
```
FORMAT_HEADING_ITEM(C2 + " - ", 20, 0, 0) +
FORMAT_HEADING_ITEM(C2.DESCRIPTION, 40, 1, 0) +
FORMAT_HEADING_ITEM(C2.<Default>, 20, 1, 0)
```

---

#### GET_ASSGNS()

**Purpose:** Returns assignment data for activity

**Data Type:** Character

**Syntax:** `GET_ASSGNS(<fieldname1>[|<fieldname2>...])`

**Format:** Fields separated by commas, records by semicolons

**Example:** `GET_ASSGNS("RES_ID|RES_LEVEL")` returns "ENG,1.00;TECH.MARY,2.00"

**NOTE:** Remove table portion of fieldname when using this function

---

#### GET_CHILDREN()

**Purpose:** Returns list of fields from immediate children of hierarchical record

**Data Type:** Character

**Syntax:** `GET_CHILDREN(<FieldList>)`

**Format:** Pipe-delimited field list, comma-separated values, semicolon-separated records

**Example:** `GET_CHILDREN("ACT_ID|DESCRIPTION")` for activity "1.01" returns:
```
"1.01.01,First child of 1.01;1.01.02,Second child of 1.01"
```

---

#### GET_COSTS()

**Purpose:** Returns cost records for activity

**Data Type:** Character

**Syntax:** `GET_COSTS(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_COSTS("ACWP_CST|ACWP_QTY")` returns "1000.00,1.00;4000.00,2.00"

---

#### GET_FIELD()

**Purpose:** Displays data from other tables or other records in same table

**Data Type:** Character

**Syntax:** `GET_FIELD(<TableType>, <UniqueID>, <FieldName>)`

**Parameters:**
- TableType: "Activity", "Resource", "Calendar", "B1"/"B2"/"B3", "ProjDir", "CodeDir", "ResDir", "CalDir"
- UniqueID: Lookup key (activity ID, project name, etc.)
- FieldName: Field to return (MUST be in quotes)

**Example:** `GET_FIELD("C2", PARENT(C2), "DESCRIPTION")`

If current activity has code 2 value "1.2.1.3", returns "1.2.1. System Engineering"

**NOTE:** "Resource" pulls from "Resource Data" table

---

#### GET_FIRST_RECORD_IN_SUMMARY()

**Purpose:** Returns field value; for grouping summary row, returns value from first child

**Data Type:** Same as input field

**Syntax:** `GET_FIRST_RECORD_IN_SUMMARY(<FieldName>)`

**Use Case:** Allows filter on break column value rather than summary value in grouping rows

**Example:**
```
GET_FIRST_RECORD_IN_SUMMARY("TOTALFLOAT") > |2d|
```
Works correctly for both detail and summary rows in grouped view

---

#### GET_NOTE()

**Purpose:** Returns activity, resource, or code note

**Data Type:** Character

**Syntax:** `GET_NOTE(<category>)`

**Parameters:**
- category: Note category name (optional, defaults to default category)

**Examples:**
- `GET_NOTE("Document")` returns Document category note
- `GET_NOTE("")` returns default category note

---

#### GET_PREDS()

**Purpose:** Returns predecessor data for activity

**Data Type:** Character

**Syntax:** `GET_PREDS(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_PREDS("PRED_ACT_ID|REL_TYPE")` returns "A100,Finish to Start"

---

#### GET_RELATED()

**Purpose:** Generic function returning fields from child collection

**Data Type:** String

**Syntax:** `GET_RELATED(<CollectionType>, <FieldList>)`

**CollectionType (Main Collection = Activity):**
- "A01", "A02", "A03" - Selected Baseline Activity
- "ASG" - Assignment collection
- "CST" - Cost collection
- "PRD" - Predecessor collection
- "RSK" - Risk collection
- "SUB" - Subproject collection
- "SUC" - Successor collection
- "U01", "U02", "U03" - Selected Baseline Usage
- "USE" - Usage collection

**CollectionType (Main Collection = Baseline Activity):**
- "SUB" - Subproject collection
- "USE" or "BSU" - Baseline usage

**CollectionType (Main Collection = Resource Data):**
- "PSU" - Project Summary Usage
- "RSL" - Resource Escalations
- "SKL" - All skill assignments for resource
- "SKR" - All skill assignments for skill

**FieldList:** Pipe-delimited fields; comma-separated values, semicolon-separated records

**Example:** `GET_RELATED("A01", "ACT_ID|ESDATE")` returns "A,12Oct04"

---

#### GET_RELATED_COUNT()

**Purpose:** Returns record count for related collection

**Data Type:** String

**Syntax:** `GET_RELATED_COUNT(<RELATED_COLLECTION_IDENTIFIER>, <FILTER_ON_RELATED_COLLECTION>)`

**Identifiers (Main Collection = Activity):**
- "ASG", "CHILD", "CST", "PRD", "RSK", "STP", "SUB", "SUC", "USE"

**Identifiers (Main Collection = Baseline Activity):**
- "CHILD", "USE"

**Examples:**
- `GET_RELATED_COUNT("CHILD")` - Count all children for subproject activity
- `GET_RELATED_COUNT("PRD", "REL_LAG = |0|")` - Count predecessors with zero lag

---

#### GET_RISKS()

**Purpose:** Returns risk data for activity

**Data Type:** Character

**Syntax:** `GET_RISKS(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_RISKS("ESDATE1")` returns each early start date from risk analysis

---

#### GET_SUCCS()

**Purpose:** Returns successor data for activity

**Data Type:** Character

**Syntax:** `GET_SUCCS(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_SUCCS("SUCC_ACT_ID|REL_TYPE")` returns "A200,Finish to Start"

---

#### GET_USAGES()

**Purpose:** Returns usage records for activity

**Data Type:** Character

**Syntax:** `GET_USAGES(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_USAGES("RES_ID|RES_USED")` returns "ENG,24"

---

#### GO_MONTH()

**Purpose:** Returns date N months before/after specified date

**Data Type:** Date

**Syntax:** `GO_MONTH(<date>, <integer>)`

**Example:** `GO_MONTH({04OCT04}, -2)` returns {04AUG04}

---

#### HAS_NOTE()

**Purpose:** Returns logical indicating if note attached

**Data Type:** Logical

**Syntax:** `HAS_NOTE(<string>)`

**Parameters:**
- string: Optional category name

**Examples:**
- `HAS_NOTE()` returns TRUE if any note attached
- `HAS_NOTE("Scope")` returns TRUE if note in "Scope" category attached

---

#### IIF()

**Purpose:** Conditional processing (if-then-else)

**Data Type:** Any (matches return value type)

**Syntax:** `IIF(<logicexp>, <iftrue>, <iffalse>)`

**Operation:** If logicexp is TRUE, return iftrue; else return iffalse

**Example:** `IIF(ESDATE > {01JUL01}, "Underway", "Planned")` returns "Planned" if ESDATE is {19JUN01}

---

#### INLIST()

**Purpose:** Returns logical indicating if value in list

**Data Type:** Logical

**Syntax:** `INLIST(<search>, <value1>, <value2>, ...)`

**Examples:**
- `INLIST(MONTH(ESDATE), 1, 4, 7, 10)` returns FALSE if month is June
- `INLIST("CHRIS", GET_ASSGNS("LOCAL(RES_ID)"))` returns TRUE for activities where CHRIS assigned

**NOTE:** Character argument interpreted as comma-delimited list

---

#### INSTR()

**Purpose:** Returns position of first occurrence of string within another

**Data Type:** Integer

**Syntax:** `INSTR(<start>, <string1>, <string2>)`

**Parameters:**
- start: Starting position for search
- string1: String being searched
- string2: String being sought

**Example:** `INSTR(1, "SITE COORDINATION AND DESIGN", "COORDINATION")` returns `6`

**NOTE:** Case sensitive

---

#### LEFT()

**Purpose:** Returns leftmost N characters

**Data Type:** Character

**Syntax:** `LEFT(<string>, <int>)`

**Example:** `LEFT("SITE COORDINATION AND DESIGN", 4)` returns "SITE"

---

#### LEN()

**Purpose:** Returns length of variable or constant

**Data Type:** Integer

**Syntax:** `LEN(<data>)`

**Example:**
- `LEN("Dig Hole")` returns `8`
- `LEN(DESCRIPTION)` where DESCRIPTION="EXCAVATION" returns `10` (not field width)

---

#### LEVEL()

**Purpose:** Returns hierarchical level of ID or code

**Data Type:** Integer

**Syntax:** `LEVEL(<ID>)`

**Example:** `LEVEL(C1)` returns level of C1 in code structure

---

#### LOCAL()

**Purpose:** Returns local portion of ID/code (not shared with siblings)

**Data Type:** Character

**Syntax:** `LOCAL(<ID>, <level>)`

**Parameters:**
- ID: Character variable or constant
- level: Optional - specific level for local portion

**Example:** `LOCAL(C1)` returns rightmost portion of C1

---

#### LOWER()

**Purpose:** Converts uppercase to lowercase

**Data Type:** Character

**Syntax:** `LOWER(<string>)`

**Example:** `LOWER("Dig Hole")` returns "dig hole"

---

#### LTRIM()

**Purpose:** Trims leading spaces

**Data Type:** Character

**Syntax:** `LTRIM(<string>)`

---

*Comprehensive calculated fields reference included. For additional functions (MAX, MIN, MID, MONTH, NEWLINE, NUMBER_FORMAT, OCCURS, PARENT, RECORD_NUMBER, RIGHT, ROUND, SPACE, SQRT, STR, STRTRAN, STUFF, SUBSTR, TIMENOW, TRIM, UPPER, USER_ID, VAL, YEAR), refer to full DeltekOpenPlanDeveloperGuide.md.*

---

*Enhanced with comprehensive VBA examples, detailed Import/Export Transfer.dat scripting, and Calculated Fields documentation.*
