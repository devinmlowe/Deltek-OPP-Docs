# Deltek Open Plan® Developer's Guide

Welcome to the Deltek Open Plan Developer's Guide. This documentation provides comprehensive information about automating and extending Deltek Open Plan through OLE Automation.

## Quick Start

New to Open Plan automation? Start here:

1. **[Overview](getting-started/overview.md)** - Introduction to Open Plan development
2. **[OLE Automation Introduction](getting-started/ole-automation-intro.md)** - Learn about OLE automation
3. **[Accessing Objects](getting-started/accessing-objects.md)** - How to work with Open Plan objects

## Core Documentation

### Getting Started
- [Overview](getting-started/overview.md) - About Open Plan and its integration capabilities
- [OLE Automation Introduction](getting-started/ole-automation-intro.md) - Understanding OLE automation
- [Accessing OLE Automation Objects](getting-started/accessing-objects.md) - Syntax and examples for accessing objects

### Object Hierarchy
- [Object Hierarchy Overview](object-hierarchy/README.md) - Visual representations of object relationships
- [Application Hierarchy](object-hierarchy/README.md#application-hierarchy) - OPCreateApplication33 and related objects
- [Project Hierarchy](object-hierarchy/README.md#project-hierarchy) - Projects, activities, and resources
- [File Cabinet Hierarchy](object-hierarchy/README.md#file-cabinet-hierarchy) - Explorer and file management objects

### Object Reference
- **[Complete Objects Reference](objects/README.md)** - All available objects

#### Most Common Objects
- [OPCreateApplication33](objects/opcreateapplication33.md) - Main application object
- [OPProject](objects/opproject.md) - Project object
- [OPActivity](objects/opactivity.md) - Activity object
- [OPResource](objects/opresource.md) - Resource file object
- [OPCalendar](objects/opcalendar.md) - Calendar file object
- [Collections](objects/collections.md) - Working with collections

### Properties and Methods
- [Properties Reference](properties/README.md) - All available properties
- [Methods Reference](methods/README.md) - All available methods

## Key Concepts

### Objects and Collections

Open Plan exposes its functionality through objects. The main object types are:

- **Application** - The [OPCreateApplication33](objects/opcreateapplication33.md) object
- **Projects** - [OPProject](objects/opproject.md) objects contain project data
- **Activities** - [OPActivity](objects/opactivity.md) objects represent tasks
- **Resources** - [OPResourceRecord](objects/opresourcerecord.md) and related objects
- **Views** - [OPView](objects/opview.md) objects for barcharts, networks, etc.

### Properties

Properties are attributes that you can get or set on objects. For example:

```vb
'Get a property
Dim actID As String
actID = activity.ID

'Set a property
activity.Description = "New Description"
activity.Duration = "10D"
```

See [Properties Reference](properties/README.md) for complete details.

### Methods

Methods are actions that objects can perform. For example:

```vb
'Perform time analysis
project.TimeAnalyze

'Add an activity
Dim newAct As Object
Set newAct = activities.Add("ACT100", "New Activity")

'Save the project
project.Save
```

See [Methods Reference](methods/README.md) for complete details.

## Common Tasks

### Opening and Working with Projects

```vb
'Create application object
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")

'Login (required for Open Plan 8.x+)
OPApp.Login

'Open a project
Dim proj As Object
Set proj = OPApp.FileOpen("C:\Projects\MyProject.opp")

'Perform time analysis
proj.TimeAnalyze

'Save and close
proj.Save
proj.Shut
```

### Working with Activities

```vb
'Get activities collection
Dim acts As Object
Set acts = proj.Activities

'Loop through activities
Dim act As Object
Dim i As Long
For i = 1 To acts.Count
    Set act = acts.Item(i)
    Debug.Print act.ID & ": " & act.Description
Next i

'Add a new activity
Set act = acts.Add("ACT100", "New Activity")
act.Duration = "5D"
act.Calendar = "STANDARD"
```

### Working with Resources

```vb
'Get project resources
Dim resources As Object
Set resources = proj.Resources

'Loop through resources
Dim res As Object
For i = 1 To resources.Count
    Set res = resources.Item(i)
    Debug.Print res.ID & ": " & res.Description
Next i
```

### Working with Relationships

```vb
'Get activity
Set act = acts.Item("ACT100")

'Add a predecessor
Dim preds As Object
Set preds = act.Predecessors

Dim pred As Object
Set pred = preds.Add("ACT050")
pred.RelationshipType = "FS"
pred.Lag = "2D"
```

## Important Notes for Open Plan 8.x and Later

### Login Requirement

Open Plan 8.x introduced a security feature requiring login. All automation applications **must** call the Login method:

```vb
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")
OPApp.Login 'Required!
```

See [Login Method](methods/login.md) for details.

### Unsupported Objects, Properties, and Methods

The following are no longer supported in Open Plan 8.x:

**Objects:**
- OPFCProcesses
- OPFCStructures

**Properties:**
- OPPFormat
- ResidueFile

**Methods:**
- FileOpenSpecialOpenPlan
- OpenSpecialODBC
- GlobalDir
- Rights
- Processes
- Structures

See [Migration Guide](getting-started/overview.md#modifying-existing-applications-for-open-plan-8x) for more information.

## Documentation Structure

This developer guide is organized as follows:

```
developer-guide/
├── index.md (this file)
├── getting-started/
│   ├── overview.md
│   ├── ole-automation-intro.md
│   └── accessing-objects.md
├── object-hierarchy/
│   └── README.md
├── objects/
│   ├── README.md
│   ├── opcreateapplication33.md
│   ├── opproject.md
│   ├── opactivity.md
│   └── [other object files...]
├── properties/
│   └── README.md
└── methods/
    └── README.md
```

## Additional Resources

- [Original Developer Guide](../DeltekOpenPlanDeveloperGuide.md) - Complete original documentation
- Sample code and examples (see methods documentation)
- Deltek Support - For technical support

## Version Information

This documentation is based on:
- **Deltek Open Plan Developer's Guide**
- **Last Updated:** December 20, 2020
- **Applies to:** Open Plan 8.x and later

---

## Navigation

- **[Getting Started →](getting-started/overview.md)**
- **[Object Reference →](objects/README.md)**
- **[Properties Reference →](properties/README.md)**
- **[Methods Reference →](methods/README.md)**
