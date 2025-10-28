# Methods Reference

This section will contain detailed reference information for all Open Plan OLE Automation methods.

## Overview

Methods are actions that Open Plan objects can perform. Each method has:

- **Name** - The method identifier
- **Applies To** - Which objects have this method
- **Description** - What the method does
- **Parameters** - Input values required
- **Return Value** - What the method returns (if anything)
- **Syntax** - How to call the method in code

## Method Categories

### Application Methods
Methods on the [OPCreateApplication33](../objects/opcreateapplication33.md) object:
- FileOpen - Open a project
- FileNew - Create a new project
- Login - Login to Open Plan
- Version - Get version information
- And more...

See [OPCreateApplication33 Methods](../objects/opcreateapplication33.md#methods) for the complete list.

### Project Methods
Methods on the [OPProject](../objects/opproject.md) object:
- TimeAnalyze - Perform time analysis
- ResourceSchedule - Perform resource scheduling
- CalculateCost - Calculate costs
- Save - Save the project
- Activities - Get the activities collection
- And many more...

See [OPProject Methods](../objects/opproject.md#methods) for the complete list.

### Activity Methods
Methods on the [OPActivity](../objects/opactivity.md) object:
- GetField - Get a field value
- SetField - Set a field value
- ActivityResources - Get activity resources
- Predecessors - Get predecessors
- Remove - Remove the activity
- And more...

See [OPActivity Methods](../objects/opactivity.md#methods) for the complete list.

### Collection Methods
Methods common to most collections:
- **Add** - Add a new item
- **Item** - Get an item by index or key
- **Remove** - Remove an item
- **Fields** - Get field information

## Common Method Patterns

### Methods with No Parameters

```vb
'Perform time analysis
project.TimeAnalyze

'Minimize a window
MyCodeView.Minimize

'Save a project
project.Save
```

### Methods with One or More Parameters

```vb
'Resource scheduling with parameters
project.ResourceSchedule "Time Limited", True

'Open a file
Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")

'Add a predecessor with ID
Set pred = predecessors.Add("ACT050")
```

### Methods that Return a Value

```vb
'Get a field value
Dim suppressFlag As Variant
suppressFlag = resource.GetField("NO_LIST")

'Get version
Dim version As String
version = OPApp.Version

'Check if selected
Dim isSelected As Boolean
isSelected = activity.Selected
```

### Methods that Return an Object

```vb
'Get a calendar
Dim cal As Object
Set cal = OPApp.Calendars.Item("MYCAL")

'Get activities collection
Dim acts As Object
Set acts = project.Activities

'Get active project
Dim proj As Object
Set proj = OPApp.ActiveProject
```

## Analysis and Scheduling Methods

### Time Analysis
```vb
'Perform time analysis on a project
project.TimeAnalyze
```

### Resource Scheduling
```vb
'Time-limited scheduling
project.ResourceSchedule "Time Limited", True

'Resource-limited scheduling
project.ResourceSchedule "Resource Limited", False
```

### Cost Calculation
```vb
'Calculate costs
project.CalculateCost

'Extended cost calculation with parameters
project.CalculateCostEx params
```

### Risk Analysis
```vb
'Perform risk analysis
project.RiskAnalyze
```

### Automatic Progress
```vb
'Perform automatic progress
project.AutoProgress

'Extended automatic progress
project.AutoProgressEx params
```

## Data Access Methods

### GetField and SetField
```vb
'Get a single field value
Dim value As Variant
value = activity.GetField("USER1")

'Set a single field value
activity.SetField "USER1", "New Value"

'Get multiple fields
Dim fieldNames(2) As String
fieldNames(0) = "ACT_ID"
fieldNames(1) = "DESCRIPTN"
fieldNames(2) = "ORIG_DUR"

Dim values As Variant
values = activity.GetFields(fieldNames)
```

### GetCurrentFields and SetCurrentFields
```vb
'Get current field set values
Dim currentValues As Variant
currentValues = activity.GetCurrentFields

'Set current field set values
activity.SetCurrentFields valueArray
```

## Collection Methods

### Add Method
```vb
'Add an activity
Dim newAct As Object
Set newAct = activities.Add("ACT100", "New Activity")

'Add a code record
Dim newCode As Object
Set newCode = codes.Add("CODE01")

'Add a predecessor
Dim pred As Object
Set pred = predecessors.Add("ACT050")
```

### Item Method
```vb
'Get by numeric index
Set act = activities.Item(1)

'Get by string key/ID
Set act = activities.Item("ACT001")

'Get by name
Set cal = calendars.Item("STANDARD")
```

### Remove Method
```vb
'Remove an item by object
activity.Remove

'Remove from collection by index
activities.Remove 1

'Remove by ID
activities.Remove "ACT001"
```

## File Operations

### Opening and Closing
```vb
'Open a project
Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")

'Close with save
proj.Shut

'Close without save
proj.ShutWOSave
```

### Saving
```vb
'Save
project.Save

'Save As
project.SaveAs "C:\Projects\NewName.opp"

'Create backup
project.CreateBackup "C:\Backups\Backup.opp"
```

### Import/Export
```vb
'General export
project.GeneralExport exportParams

'General import
project.GeneralImport importParams
```

## Window Management Methods

```vb
'Show window
OPApp.Show

'Hide window
window.Conceal

'Maximize
window.Maximize

'Minimize
window.Minimize

'Restore
window.Restore

'Tile windows
OPApp.WindowTile
OPApp.WindowTileVertical
```

## Baseline Methods

```vb
'Create a baseline
project.CreateBaseline "Baseline1"

'Update a baseline
project.UpdateBaseline "Baseline1"

'Delete a baseline
project.DeleteBaseline "Baseline1"
```

## Important Methods for Open Plan 8.x

### Login Method (Required)

The Login method is **required** for Open Plan 8.x and later:

```vb
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")
OPApp.Login 'Must call before other operations
```

## See Also

- [Objects Reference](../objects/README.md)
- [Properties Reference](../properties/README.md)
- [Getting Started](../getting-started/overview.md)
- [Original Developer Guide](../../DeltekOpenPlanDeveloperGuide.md) - For detailed method descriptions
