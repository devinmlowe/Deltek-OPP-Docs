# Open Plan Object Hierarchy

This section provides visual and textual representations of how objects are organized in the Open Plan OLE Automation object model.

## Understanding the Hierarchy

The Open Plan object hierarchy shows the relationships between objects. Objects at higher levels contain or provide access to objects at lower levels through:

- **Properties** - Direct access to child objects
- **Methods** - Methods that return collections or objects
- **Collections** - Groups of related objects

## Main Object Hierarchies

### Application Hierarchy

The [OPCreateApplication33](../objects/opcreateapplication33.md) object is the root of the application hierarchy:

```
OPCreateApplication33 (Application)
├── Projects (OPProjects collection)
│   └── Project (OPProject object)
│       ├── Activities (OPActivities collection)
│       │   └── Activity (OPActivity object)
│       ├── Resources (OPProjectResources collection)
│       ├── Views (OPViews collection)
│       ├── Barcharts (OPBarcharts collection)
│       ├── Networks (OPNetworks collection)
│       ├── Spreadsheets (OPSpreadsheets collection)
│       ├── Graphs (OPGraphs collection)
│       └── CodeFiles (OPProjectCodes collection)
├── Resources (OPResources collection)
│   └── Resource (OPResource object)
│       └── ResourceRecords (collection)
│           └── ResourceRecord (OPResourceRecord object)
├── Calendars (OPCalendars collection)
│   └── Calendar (OPCalendar object)
│       └── CalendarRecords (collection)
│           └── CalendarRecord (OPCalendarRecord object)
├── Codes (OPCodes collection)
│   └── Code (OPCode object)
│       └── CodeRecords (collection)
│           └── CodeRecord (OPCodeRecord object)
├── FileCabinet (OPFileCabinet object)
│   ├── Projects (OPFCProjects collection)
│   ├── Resources (OPFCResources collection)
│   ├── Calendars (OPFCCalendars collection)
│   ├── Codes (OPFCCodes collection)
│   ├── Barcharts (OPFCBarcharts collection)
│   ├── Networks (OPFCNetworks collection)
│   ├── Spreadsheets (OPFCSpreadsheets collection)
│   └── Graphs (OPFCGraphs collection)
├── ReportingCalendars (OPReportingCalendars collection)
└── WebWindows (OPWebWindows collection)
```

### Project Hierarchy

The [OPProject](../objects/opproject.md) object contains:

```
OPProject
├── Activities (OPActivities collection)
│   └── Activity (OPActivity object)
│       ├── Predecessors (OPPredecessors collection)
│       ├── Assignments (OPAssignments collection)
│       ├── ActivityResources (OPActivityResources collection)
│       ├── ResourceUsages (OPResourceUsages collection)
│       ├── Notes (OPNotes collection)
│       └── Steps (OPSteps collection)
├── Resources (OPProjectResources collection)
│   └── ProjectResource (OPProjectResource object)
│       ├── Availabilities (OPAvailabilities collection)
│       └── Notes (OPNotes collection)
├── CodeFiles (OPProjectCodes collection)
│   └── ProjectCode (OPProjectCode object)
│       └── CodeRecords (collection)
│           └── CodeRecord (OPCodeRecord object)
├── Costs (OPCosts collection)
├── BaselinesList (OPBaselinesList collection)
│   └── Baselines (OPBaselines collection)
│       └── Activity (OPActivity object - baseline)
├── Filters (OPFilters collection)
├── Sorts (OPSorts collection)
├── GlobalEdits (OPGlobalEdits collection)
├── BatchGlobalEdits (OPBatchGlobalEdits collection)
├── CalculatedFields (OPCalculatedFields collection)
├── Rollups (OPRollups collection)
├── Categories (OPCategories collection)
├── Notes (OPNotes collection)
├── Views (OPViews collection)
├── Barcharts (OPBarcharts collection)
├── Networks (OPNetworks collection)
├── Spreadsheets (OPSpreadsheets collection)
└── Graphs (OPGraphs collection)
```

### Resource Hierarchy

```
OPResource (Resource File)
├── ResourceRecords (collection)
│   └── ResourceRecord (OPResourceRecord object)
│       ├── Availabilities (OPAvailabilities collection)
│       │   └── Availability (OPAvailability object)
│       ├── Skills (OPSkills collection)
│       └── Notes (OPNotes collection)
├── Filters (OPFilters collection)
├── Sorts (OPSorts collection)
├── GlobalEdits (OPGlobalEdits collection)
├── CalculatedFields (OPCalculatedFields collection)
├── Rollups (OPRollups collection)
└── Categories (OPCategories collection)
```

### Calendar Hierarchy

```
OPCalendar (Calendar File)
└── CalendarRecords (collection)
    └── CalendarRecord (OPCalendarRecord object)
        ├── Holidays (OPHolidays collection)
        │   └── Date (OPDate object)
        ├── ExtraWorkDays (OPExtraWorkDays collection)
        │   └── Date (OPDate object)
        └── StandardDays (7 OPStandardDay objects)
            └── Shifts (OPShifts collection)
                └── Shift (OPShift object)
```

### Code Hierarchy

```
OPCode (Code File)
└── CodeRecords (collection)
    └── CodeRecord (OPCodeRecord object)
        ├── Notes (OPNotes collection)
        ├── Categories (OPCategories collection)
        └── Rollups (OPRollups collection)
```

### File Cabinet Hierarchy

```
OPFileCabinet (Explorer)
├── Projects (OPFCProjects collection)
│   └── Icon (OPIcon object)
├── Resources (OPFCResources collection)
│   └── Icon (OPIcon object)
├── Calendars (OPFCCalendars collection)
│   └── Icon (OPIcon object)
├── Codes (OPFCCodes collection)
│   └── Icon (OPIcon object)
├── Barcharts (OPFCBarcharts collection)
│   └── FCView (OPFCView object)
├── Networks (OPFCNetworks collection)
│   └── FCView (OPFCView object)
├── Spreadsheets (OPFCSpreadsheets collection)
│   └── FCView (OPFCView object)
├── Graphs (OPFCGraphs collection)
│   └── FCView (OPFCView object)
└── Views (OPFCViews collection)
    └── FCView (OPFCView object)
```

## How to Navigate the Hierarchy

### Top-Down Navigation

Start from the application and navigate down:

```vb
'Get application
Dim OPApp As Object
Set OPApp = CreateObject("opp.application")
OPApp.Login

'Get projects collection
Dim projects As Object
Set projects = OPApp.Projects

'Get a specific project
Dim proj As Object
Set proj = projects.Item(1)

'Get activities collection
Dim activities As Object
Set activities = proj.Activities

'Get a specific activity
Dim act As Object
Set act = activities.Item("ACT001")

'Get predecessors
Dim preds As Object
Set preds = act.Predecessors
```

### Direct Access

Some objects can be accessed directly:

```vb
'Get active project
Dim activeProj As Object
Set activeProj = OPApp.ActiveProject

'Get active view
Dim activeView As Object
Set activeView = OPApp.ActiveView
```

## Object Relationships

### Parent-Child Relationships

- A **Project** contains **Activities**
- An **Activity** contains **Predecessors**, **Assignments**, etc.
- A **Calendar** contains **Calendar Records**
- A **Code File** contains **Code Records**

### Reference Relationships

Some objects reference others:

- An **Activity** has an **ActivityResource** which references a **ResourceRecord**
- An **Assignment** references both an **Activity** and a **Resource**
- A **Predecessor** references two **Activities** (predecessor and successor)

## Collection vs. Object

Understanding the difference:

- **Collection** - A group of objects (e.g., OPActivities)
  - Has Count property
  - Has Item method to get individual objects
  - Has Add method to create new objects

- **Object** - An individual item (e.g., OPActivity)
  - Has properties like ID, Description, Duration
  - Has methods like GetField, SetField, Remove

```vb
'Collection
Dim activities As Object 'OPActivities collection
Set activities = proj.Activities
Debug.Print activities.Count 'Number of activities

'Object
Dim activity As Object 'OPActivity object
Set activity = activities.Item(1)
Debug.Print activity.ID 'Activity ID
```

## See Also

- [Objects Reference](../objects/README.md)
- [Getting Started](../getting-started/overview.md)
- [OPCreateApplication33](../objects/opcreateapplication33.md)
- [OPProject](../objects/opproject.md)
- [OPActivity](../objects/opactivity.md)
