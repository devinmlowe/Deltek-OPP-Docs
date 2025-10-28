# OPActivity Object

## Description

When retrieved from the [OPActivities collection](collections.md#opactivities-collection), the OPActivity object contains information about a single activity within an open project. When retrieved from the [OPBaselines collection](opbaselines.md), the OPActivity object contains information about a single baseline activity within a saved baseline belonging to an open project.

## How to Get This Object

```vb
'Via the Activities collection
Dim proj As Object
Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")

Dim acts As Object
Set acts = proj.Activities

'Get by index
Dim act As Object
Set act = acts.Item(1)

'Get by ID
Set act = acts.Item("ACT001")

'Or add a new activity
Set act = acts.Add("ACT100", "New Activity Description")
```

## Properties

### Identification and Status
- [ID](../properties/id.md) - Unique activity identifier
- [Description](../properties/description.md) - Activity description
- [Type](../properties/type.md) - Activity type
- [ActivityLogicFlag](../properties/activitylogicflag.md) - Start/finish activity indicator
- [ComputedStatus](../properties/computedstatus.md) - Computed activity status (Planned, In Progress, Complete)
- [KeyActivityStatus](../properties/keyactivitystatus.md) - Key activity status

### Date Properties
- [ActualFinishDate](../properties/actualfinishdate.md) - Actual finish date
- [ActualStartDate](../properties/actualstartdate.md) - Actual start date
- [BaselineFinish](../properties/baselinefinish.md) - Baseline finish date
- [BaselineStart](../properties/baselinestart.md) - Baseline start date
- [EarliestFeasible](../properties/earliestfeasible.md) - Earliest feasible start date
- [EarlyFinish](../properties/earlyfinish.md) - Early finish date
- [EarlyStart](../properties/earlystart.md) - Early start date
- [ExpectedFinish](../properties/expectedfinish.md) - Expected finish date
- [LateFinish](../properties/latefinish.md) - Late finish date
- [LateStart](../properties/latestart.md) - Late start date
- [ScheduledFinish](../properties/scheduledfinish.md) - Scheduled finish date
- [ScheduledStart](../properties/scheduledstart.md) - Scheduled start date
- [TargetFinish](../properties/targetfinish.md) - Target finish date
- [TargetFinishType](../properties/targetfinishtype.md) - Target finish type
- [TargetStart](../properties/targetstart.md) - Target start date
- [TargetStartType](../properties/targetstarttype.md) - Target start type

### Duration Properties
- [Calendar](../properties/calendar.md) - Activity calendar
- [ComputedRemainingDuration](../properties/computedremainingduration.md) - Computed remaining duration
- [Duration](../properties/duration.md) - Original duration
- [DurationDistShape](../properties/durationdistshape.md) - Duration distribution shape (for risk analysis)
- [MaxDuration](../properties/maxduration.md) - Maximum duration (risk analysis)
- [MaxNoSplits](../properties/maxnosplits.md) - Maximum number of splits
- [MinSplitLength](../properties/minsplitlength.md) - Minimum split length
- [OptimisticDuration](../properties/optimisticduration.md) - Optimistic duration (risk analysis)
- [PessimisticDuration](../properties/pessimisticduration.md) - Pessimistic duration (risk analysis)
- [ScheduleDuration](../properties/scheduleduration.md) - Scheduled duration

### Float and Critical Path
- [Critical](../properties/critical.md) - Critical path indicator
- [CriticalIndex](../properties/criticalindex.md) - Critical index (risk analysis)
- [EarlyFinishStdDev](../properties/earlyfinishstddev.md) - Early finish standard deviation
- [EarlyStartStdDev](../properties/earlystartst ddev.md) - Early start standard deviation
- [FinishFreeFloat](../properties/finishfreefloa t.md) - Finish free float
- [FinishTotalFloat](../properties/finishtotalfloat.md) - Finish total float
- [FreeFloat](../properties/freefloat.md) - Free float
- [FreeFloatStdDev](../properties/freefloatstddev.md) - Free float standard deviation
- [LateFinishStdDev](../properties/latefinishstddev.md) - Late finish standard deviation
- [LateStartStdDev](../properties/latestartstddev.md) - Late start standard deviation
- [ScheduleFloat](../properties/schedulefloat.md) - Schedule float
- [TotalFloat](../properties/totalfloat.md) - Total float
- [TotalFloatStdDev](../properties/totalfloatstddev.md) - Total float standard deviation

### Progress Properties
- [PercentComplete](../properties/percentcomplete.md) - Percent complete
- [ProgressFlag](../properties/progressflag.md) - Progress flag
- [ProgressValue](../properties/progressvalue.md) - Progress value

### Cost Properties
- [BudgetCost](../properties/budgetcost.md) - Budget cost
- [CostToDate](../properties/costtodate.md) - Actual cost to date
- [TotalResourceCost](../properties/totalresourcecost.md) - Total resource cost

### Resource Properties
- [DelayingResource](../properties/delayingresource.md) - ID of delaying resource
- [FirstUsage](../properties/firstusage.md) - First resource usage
- [ResourceScheduleType](../properties/resourcescheduletype.md) - Resource schedule type
- [SuppressRequirements](../properties/suppressrequirements.md) - Suppress requirements flag

### Other Properties
- [RefreshData](../properties/refreshdata.md) - Refresh data flag
- [ScheduleActions](../properties/scheduleactions.md) - Schedule actions
- [SubprojectFilename](../properties/subprojectfilename.md) - Subproject filename

## Methods

### Collections Access
- **[ActivityResources](../methods/activityresources.md)** - Access activity resources
- **[Assignments](../methods/assignments.md)** - Access resource [assignments](opassignment.md)
- **[Notes](../methods/notes.md)** - Access activity [notes](opnote.md)
- **[Predecessors](../methods/predecessors.md)** - Access [predecessor](oppredecessor.md) relationships
- **[ResourceUsages](../methods/resourceusages.md)** - Access resource [usages](opresourceusage.md)
- **[Steps](../methods/steps.md)** - Access activity steps

### Data Operations
- **[GetCurrentFields](../methods/getcurrentfields.md)** - Get current field values
- **[GetField](../methods/getfield.md)** - Get a specific field value
- **[GetFields](../methods/getfields.md)** - Get multiple field values
- **[SetCurrentFields](../methods/setcurrentfields.md)** - Set current field values
- **[SetField](../methods/setfield.md)** - Set a specific field value

### Other Methods
- **[IsExternalOpen](../methods/isexternalopen.md)** - Check if external project is open
- **[Remove](../methods/remove.md)** - Remove the activity
- **[Selected](../methods/selected.md)** - Check if activity is selected

## Usage Examples

### Getting Activity Information

```vb
Dim proj As Object
Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")

Dim acts As Object
Set acts = proj.Activities

Dim act As Object
Set act = acts.Item("ACT001")

'Get basic info
Debug.Print "ID: " & act.ID
Debug.Print "Description: " & act.Description
Debug.Print "Duration: " & act.Duration
Debug.Print "Early Start: " & act.EarlyStart
Debug.Print "Early Finish: " & act.EarlyFinish
Debug.Print "Status: " & act.ComputedStatus
```

### Modifying Activity Properties

```vb
'Set dates and duration
act.ActualStartDate = #1/15/2024#
act.PercentComplete = 50
act.Duration = "10D"

'Set description
act.Description = "Updated Activity Description"

'Set calendar
act.Calendar = "STANDARD"
```

### Working with Activity Resources

```vb
'Get activity resources
Dim actResources As Object
Set actResources = act.ActivityResources

'Loop through resources assigned to this activity
Dim actRes As Object
Dim i As Long
For i = 1 To actResources.Count
    Set actRes = actResources.Item(i)
    Debug.Print "Resource: " & actRes.ID
    Debug.Print "Units: " & actRes.Units
Next i
```

### Working with Predecessors

```vb
'Get predecessors
Dim preds As Object
Set preds = act.Predecessors

'Add a predecessor
Dim pred As Object
Set pred = preds.Add("ACT050")
pred.RelationshipType = "FS"
pred.Lag = "2D"

'List all predecessors
For i = 1 To preds.Count
    Set pred = preds.Item(i)
    Debug.Print "Predecessor: " & pred.ID & " " & pred.RelationshipType
Next i
```

### Adding Activity Notes

```vb
'Get notes collection
Dim notes As Object
Set notes = act.Notes

'Add a note
Dim note As Object
Set note = notes.Add("General")
note.NoteText = "This is an important activity."
```

### Using GetField and SetField

```vb
'Get a user-defined field
Dim udField As Variant
udField = act.GetField("USER1")
Debug.Print "USER1: " & udField

'Set a user-defined field
act.SetField "USER1", "New Value"

'Get multiple fields
Dim fieldNames(2) As String
fieldNames(0) = "ACT_ID"
fieldNames(1) = "DESCRIPTN"
fieldNames(2) = "ORIG_DUR"

Dim fieldValues As Variant
fieldValues = act.GetFields(fieldNames)

Dim j As Long
For j = LBound(fieldValues) To UBound(fieldValues)
    Debug.Print fieldNames(j) & ": " & fieldValues(j)
Next j
```

## See Also

- [OPActivities Collection](collections.md#opactivities-collection)
- [OPProject](opproject.md)
- [OPAssignment](opassignment.md)
- [OPPredecessor](oppredecessor.md)
- [Properties Reference](../properties/README.md)
- [Methods Reference](../methods/README.md)
