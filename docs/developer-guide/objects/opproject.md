# OPProject Object

## Description

The OPProject object contains properties and methods for an open project within the [OPProjects](collections.md#opprojects-collection) collection. This is one of the most important objects for project automation.

## How to Get This Object

```vb
'Via the FileOpen method
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")
OPApp.Login

Dim proj As Object
Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")

'Or via the ActiveProject method
Dim activeProj As Object
Set activeProj = OPApp.ActiveProject

'Or via the Projects collection
Dim projects As Object
Set projects = OPApp.Projects
Dim firstProj As Object
Set firstProj = projects.Item(1)
```

## Properties

### Project Information
- [AccessMode](../properties/accessmode.md) - Default access mode for the project
- [Calendar](../properties/calendar.md) - Calendar file used by the project
- [Client](../properties/client.md) - Client name
- [Company](../properties/company.md) - Company name
- [CurrentActivity](../properties/currentactivity.md) - Currently selected activity ID
- [DateFormat](../properties/dateformat.md) - Default date format
- [Description](../properties/description.md) - Project description
- [FileName](../properties/filename.md) - Project file name
- [FiscalCalendar](../properties/fiscalcalendar.md) - Fiscal calendar
- [Manager](../properties/manager.md) - Project manager name
- [Name](../properties/name.md) - Project name
- [Resource](../properties/resource.md) - Resource file

### Date Properties
- [EarlyFinishDate](../properties/earlyfinishdate.md) - Early finish date
- [LateFinishDate](../properties/latefinishdate.md) - Late finish date
- [ScheduleFinishDate](../properties/schedulefinishdate.md) - Scheduled finish date
- [StartDate](../properties/startdate.md) - Project start date
- [StatusDate](../properties/statusdate.md) - Status/time now date
- [TargetCost](../properties/targetcost.md) - Target cost
- [TargetFinishDate](../properties/targetfinishdate.md) - Target finish date
- [TargetFinishType](../properties/targetfinishtype.md) - Target finish type
- [TargetStartDate](../properties/targetstartdate.md) - Target start date

### Duration and Time Unit Properties
- [DefaultActivityCalendar](../properties/defaultactivitycalendar.md) - Default activity calendar
- [DefaultActivityType](../properties/defaultactivitytype.md) - Default activity type
- [DefaultDurationChar](../properties/defaultdurationchar.md) - Default duration unit character
- [DefaultDurationUnit](../properties/defaultdurationunit.md) - Default duration unit
- [DefaultRelationshipCalendar](../properties/defaultrelationshipcalendar.md) - Default relationship calendar
- [MinimumCalcDurationUnit](../properties/minimumcalcdurationunit.md) - Minimum calculation duration unit
- [MinutesPerDay](../properties/minutesperday.md) - Minutes per day
- [MinutesPerDefaultUnit](../properties/minutesperdefaultunit.md) - Minutes per default unit
- [MinutesPerMinDurUnit](../properties/minutespermindurunit.md) - Minutes per minimum duration unit
- [MinutesPerMonth](../properties/minutespermonth.md) - Minutes per month
- [MinutesPerWeek](../properties/minutesperweek.md) - Minutes per week

### Auto Progress Properties
- [AutoProgActBasedOn](../properties/autoprogactbasedon.md) - Based-on dates for auto progress
- [AutoProgActComplete](../properties/autoprogactcomplete.md) - Auto complete activities
- [AutoProgActFilter](../properties/autoprogactfilter.md) - Filter for auto progress
- [AutoProgActInProgress](../properties/autoprogactinprogress.md) - Auto set in-progress
- [AutoProgActivity](../properties/autoprogactivity.md) - Auto progress activities
- [AutoProgActProgressType](../properties/autoprogactprogresstype.md) - Progress calculation type
- [AutoProgActSetPPC](../properties/autoprogactsetppc.md) - Auto set physical percent complete
- [AutoProgResEndDate](../properties/autoprogresenddate.md) - Resource progress end date
- [AutoProgResource](../properties/autoprogresource.md) - Auto progress resources
- [AutoProgResStartDate](../properties/autoprogresstartdate.md) - Resource progress start date

### Cost Calculation Properties
- [CalcCostActual](../properties/calccostactual.md) - Calculate actual cost
- [CalcCostBasedOn](../properties/calccostbasedon.md) - Based-on dates for cost calculation
- [CalcCostBudget](../properties/calccostbudget.md) - Calculate budget cost
- [CalcCostEarnedValues](../properties/calccosteau.md) - Calculate earned values
- [CalcCostEscalated](../properties/calccostescalated.md) - Use escalated costs
- [CalcCostIncludeChildValues](../properties/calccostincludechildvalues.md) - Include child activity values
- [CalcCostRemaining](../properties/calccostremaining.md) - Calculate remaining cost

### Schedule Properties
- [AutoAnalyze](../properties/autoanalyze.md) - Auto time analysis flag
- [DefaultAuxiliaryAccessMode](../properties/defaultauxiliaryaccessmode.md) - Default auxiliary access mode
- [DefaultProjectAccessMode](../properties/defaultprojectaccessmode.md) - Default project access mode
- [HardZeroes](../properties/hardzeroes.md) - Hard zeroes option
- [InProgressPriority](../properties/inprogresspriority.md) - In-progress priority
- [MultipleEnd](../properties/multipleend.md) - Multiple end activities option
- [OutOfSeqOpt](../properties/outofseqopt.md) - Out of sequence option
- [Priority1Name](../properties/priority1name.md) - Priority 1 name
- [Priority2Name](../properties/priority2name.md) - Priority 2 name
- [Priority3Name](../properties/priority3name.md) - Priority 3 name
- [ScheduleMethod](../properties/schedulemethod.md) - Schedule method
- [ScheduleTimeUnit](../properties/scheduletimeunit.md) - Schedule time unit
- [ScheduleTimeUnitMinutes](../properties/scheduletimeunitminutes.md) - Schedule time unit in minutes
- [Smoothing](../properties/smoothing.md) - Resource smoothing option
- [StartView](../properties/startview.md) - Start view

### Window Properties
- [Height](../properties/height.md) - Window height
- [Width](../properties/width.md) - Window width
- [XPosition](../properties/xposition.md) - Window X position
- [YPosition](../properties/yposition.md) - Window Y position

## Methods

### Collections Access
- **[Activities](../methods/activities.md)** - Access the [OPActivities](opactivity.md) collection
- **[Barcharts](../methods/barcharts.md)** - Access the [OPBarcharts](collections.md#opbarcharts-collection) collection
- **[BaselinesList](../methods/baselineslist.md)** - Access the [OPBaselinesList](collections.md#opbaselineslist-collection)
- **[BatchGlobalEdits](../methods/batchglobaledits.md)** - Access batch global edits
- **[CalculatedFields](../methods/calculatedfields.md)** - Access calculated fields
- **[Categories](../methods/categories.md)** - Access categories
- **[CodeFiles](../methods/codefiles.md)** - Access code files
- **[Costs](../methods/costs.md)** - Access the [OPCosts](collections.md#opcosts-collection) collection
- **[Filters](../methods/filters.md)** - Access filters
- **[GlobalEdits](../methods/globaledits.md)** - Access global edits
- **[Graphs](../methods/graphs.md)** - Access graphs/histograms
- **[Networks](../methods/networks.md)** - Access networks
- **[Notes](../methods/notes.md)** - Access notes
- **[Resources](../methods/resources.md)** - Access project resources
- **[Rollups](../methods/rollups.md)** - Access rollups
- **[Sorts](../methods/sorts.md)** - Access sorts
- **[Spreadsheets](../methods/spreadsheets.md)** - Access spreadsheets
- **[Views](../methods/views.md)** - Access all views

### Analysis and Calculation
- **[AutoProgress](../methods/autoprogress.md)** - Perform automatic progress
- **[AutoProgressEx](../methods/autoprogressex.md)** - Perform automatic progress (extended)
- **[CalculateCost](../methods/calculatecost.md)** - Calculate costs
- **[CalculateCostEx](../methods/calculatecostex.md)** - Calculate costs (extended)
- **[ResourceSchedule](../methods/resourceschedule.md)** - Perform resource scheduling
- **[RiskAnalyze](../methods/riskanalyze.md)** - Perform risk analysis
- **[TimeAnalyze](../methods/timeanalyze.md)** - Perform time analysis

### Baseline Management
- **[CreateBaseline](../methods/createbaseline.md)** - Create a new baseline
- **[DeleteBaseline](../methods/deletebaseline.md)** - Delete a baseline
- **[UpdateBaseline](../methods/updatebaseline.md)** - Update a baseline
- **[UpdateBaselineEx](../methods/updatebaselineex.md)** - Update a baseline (extended)

### Data Operations
- **[GetField](../methods/getfield.md)** - Get a field value
- **[GetFields](../methods/getfields.md)** - Get multiple field values
- **[SetField](../methods/setfield.md)** - Set a field value
- **[GetSelectedActivitiesArray](../methods/getselectedactivitiesarray.md)** - Get selected activities

### Crosstab Operations
- **[GenerateCrosstabDates](../methods/generatecrosstabdates.md)** - Generate crosstab dates
- **[GetCrosstabDates](../methods/getcrosstabdates.md)** - Get crosstab dates
- **[GetCrosstabDatesInXML](../methods/getcrosstabdatesinxml.md)** - Get crosstab dates in XML
- **[SetCrosstabDates](../methods/setcrosstabdates.md)** - Set crosstab dates
- **[SetCrosstabDatesFromXML](../methods/setcrosstabdatesfromxml.md)** - Set crosstab dates from XML
- **[SetEarnedValueCrosstabOptions](../methods/setearnedvaluecrosstaboptions.md)** - Set earned value options
- **[SetResourceCrosstabOptions](../methods/setresourcecrosstaboptions.md)** - Set resource crosstab options

### File Operations
- **[CreateBackup](../methods/createbackup.md)** - Create a backup
- **[GeneralExport](../methods/generalexport.md)** - Export data
- **[GeneralImport](../methods/generalimport.md)** - Import data
- **[GetSharedModeLock](../methods/getsharedmodelock.md)** - Get shared mode lock status
- **[Save](../methods/save.md)** - Save the project
- **[SaveAs](../methods/saveas.md)** - Save the project with a new name
- **[SetSharedmodeLock](../methods/setsharedmodelock.md)** - Set shared mode lock
- **[Shut](../methods/shut.md)** - Close the project
- **[ShutWOSave](../methods/shutwosave.md)** - Close without saving

### Window Management
- **[CloseView](../methods/closeview.md)** - Close a view
- **[Conceal](../methods/conceal.md)** - Hide the project window
- **[Disable](../methods/disable.md)** - Disable the project window
- **[Enable](../methods/enable.md)** - Enable the project window
- **[Maximize](../methods/maximize.md)** - Maximize the project window
- **[Minimize](../methods/minimize.md)** - Minimize the project window
- **[Restore](../methods/restore.md)** - Restore the project window

## Usage Examples

### Opening and Analyzing a Project

```vb
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")
OPApp.Login

'Open the project
Dim proj As Object
Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")

'Perform time analysis
proj.TimeAnalyze

'Perform resource scheduling
proj.ResourceSchedule "Time Limited", True

'Calculate costs
proj.CalculateCost

'Save the project
proj.Save
```

### Working with Activities

```vb
'Get the activities collection
Dim acts As Object
Set acts = proj.Activities

'Count activities
Debug.Print "Total activities: " & acts.Count

'Loop through activities
Dim act As Object
Dim i As Long
For i = 1 To acts.Count
    Set act = acts.Item(i)
    Debug.Print act.ID & ": " & act.Description
Next i
```

### Auto Progress

```vb
'Configure auto progress settings
proj.AutoProgActivity = True
proj.AutoProgActBasedOn = "EARLY"
proj.AutoProgActInProgress = True
proj.AutoProgResource = True

'Perform auto progress
proj.AutoProgress
```

## See Also

- [OPCreateApplication33](opcreateapplication33.md)
- [OPActivity](opactivity.md)
- [OPActivities Collection](collections.md#opactivities-collection)
- [Properties Reference](../properties/README.md)
- [Methods Reference](../methods/README.md)
