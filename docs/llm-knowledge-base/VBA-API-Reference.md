# VBA API Reference for Deltek Open Plan®

**Purpose:** Complete VBA automation API reference with objects, properties, methods, and working examples.

**⚠️ CRITICAL:** Before using this reference, check [Critical-Warnings-and-Patterns.md](Critical-Warnings-and-Patterns.md) for fatal mistakes and required patterns.

---

## Quick Reference

### Initialize Application
```vb
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")
OPApp.Login  '⚠️ REQUIRED for Open Plan 8.x+
```

🐛 **Common Issue:** Forgetting `.Login()` causes silent failure in 8.x+

### Open & Analyze Project
```vb
Set proj = OPApp.FileOpen("C:\path\project.opp")
proj.TimeAnalyze  '⚠️ REQUIRED after activity changes
proj.ResourceSchedule "Time Limited", True
proj.CalculateCost
proj.Save  '⚠️ REQUIRED before closing
proj.Shut
```

🐛 **Common Issue:** Not calling `.TimeAnalyze()` results in stale dates

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

🐛 **Debugging Tips:**
- "Permission denied" error → Check .Login() was called
- "Object not found" → Verify file paths are absolute, not relative

✅ **Best Practice:** Always call .Login() immediately after CreateObject for 8.x+

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

🐛 **Debugging Tips:**
- Dates not updating → Call .TimeAnalyze() after modifications
- Changes not saved → Call .Save() before .Shut()
- "Calendar not found" → Check calendar file path is absolute

✅ **Best Practices:**
- Set AutoAnalyze = False before bulk changes, re-enable and call .TimeAnalyze() once at end
- Always call .Save() before .Shut() to persist changes
- Use .CreateBackup() before major operations

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

🐛 **Debugging Tips:**
- Early/Late dates not updating → Call proj.TimeAnalyze()
- "Type mismatch" on Duration → Must include unit: "10D" not "10"
- GetField returns wrong type → Use database field names (ACT_ID not "Activity ID")

✅ **Best Practices:**
- Use .GetFields() for multiple fields (faster than multiple .GetField() calls)
- Always set RefreshData = True in shared mode before reading critical data
- Check ComputedStatus instead of PercentComplete for workflow logic

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

✅ **Best Practices:**
- Apply filters before large loops to reduce iteration count
- Use .Sort before iteration for predictable ordering
- Clear filter with `acts.Filter = ""` when done

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
Set codeFiles = proj.CodeFiles
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

✅ **Note:** For detailed Import/Export scripting using Transfer.dat, see [Import-Export-Reference.md](Import-Export-Reference.md)

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

*For critical warnings, anti-patterns, and troubleshooting, see [Critical-Warnings-and-Patterns.md](Critical-Warnings-and-Patterns.md)*
*For Import/Export scripting details, see [Import-Export-Reference.md](Import-Export-Reference.md)*
*For calculated field functions, see [Calculated-Fields-Reference.md](Calculated-Fields-Reference.md)*

**Last Updated:** 2025-10-28
