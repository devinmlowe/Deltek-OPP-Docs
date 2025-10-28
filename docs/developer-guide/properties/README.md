# Properties Reference

This section will contain detailed reference information for all Open Plan OLE Automation properties.

## Overview

Properties represent the attributes of Open Plan objects that can be read (Get) or modified (Set). Each property has:

- **Name** - The property identifier
- **Applies To** - Which objects have this property
- **Description** - What the property represents
- **Access Type** - Get, Set, or Get/Set
- **Data Type** - String, Boolean, Long, Double, Date, etc.
- **Syntax** - How to use the property in code

## Property Categories

### Project Properties
Properties that apply to the [OPProject](../objects/opproject.md) object, such as:
- AccessMode
- Calendar
- DateFormat
- StartDate
- StatusDate
- And many more...

See [OPProject Properties](../objects/opproject.md#properties) for the complete list.

### Activity Properties
Properties that apply to the [OPActivity](../objects/opactivity.md) object, such as:
- ID
- Description
- Duration
- EarlyStart
- EarlyFinish
- PercentComplete
- And many more...

See [OPActivity Properties](../objects/opactivity.md#properties) for the complete list.

### Resource Properties
Properties that apply to resource objects, such as:
- Class
- Type
- Units
- UnitCost
- And more...

See [Resource Objects](../objects/README.md#resource-objects) for details.

### Collection Properties
Properties common to most collections:
- **Count** - Number of items in the collection
- **Filter** - Filter applied to the collection
- **Sort** - Sort order for the collection

### Window Properties
Properties that control window appearance:
- Height
- Width
- XPosition
- YPosition

## Common Property Patterns

### Getting a Property Value

```vb
'Get a string property
Dim actID As String
actID = activity.ID

'Get a date property
Dim startDate As Date
startDate = activity.EarlyStart

'Get a numeric property
Dim duration As Double
duration = activity.Duration

'Get a boolean property
Dim isComplete As Boolean
isComplete = (activity.PercentComplete = 100)
```

### Setting a Property Value

```vb
'Set a string property
activity.Description = "New Description"

'Set a date property
activity.ActualStartDate = #1/15/2024#

'Set a numeric property
activity.PercentComplete = 50

'Set a boolean property
project.AutoAnalyze = True
```

### Read-Only Properties

Some properties are read-only (Get only). Attempting to set them will cause an error:

```vb
'This is OK - reading a read-only property
Dim earlyStart As Date
earlyStart = activity.EarlyStart

'This will cause an error - trying to set a read-only property
'activity.EarlyStart = #1/15/2024# 'ERROR!
```

## Property Data Types

### String
Text values, such as IDs, descriptions, file names:
```vb
activity.ID = "ACT001"
activity.Description = "Task Description"
```

### Date
Date and time values:
```vb
activity.ActualStartDate = #1/15/2024#
project.StatusDate = Now()
```

### Long (Integer)
Whole numbers:
```vb
Dim count As Long
count = activities.Count
```

### Double
Decimal numbers:
```vb
activity.PercentComplete = 75.5
resource.Units = 1.5
```

### Boolean
True/False values:
```vb
project.AutoAnalyze = True
activity.Selected = False
```

## See Also

- [Objects Reference](../objects/README.md)
- [Methods Reference](../methods/README.md)
- [Getting Started](../getting-started/overview.md)
- [Original Developer Guide](../../DeltekOpenPlanDeveloperGuide.md) - For detailed property descriptions
