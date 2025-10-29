# Critical Warnings & Common Patterns

**Purpose:** This document contains essential warnings, best practices, and anti-patterns that MUST be consulted when assisting users with Deltek Open Plan® development.

---

## ⚠️ FATAL MISTAKES (Cause Silent Failures)

### VBA Automation

#### ❌ NOT Calling .Login()
**Impact:** Automation completely fails in Open Plan 8.x+ without error message
**Fix:** ALWAYS call `OPApp.Login` immediately after CreateObject

```vb
' ❌ WRONG - Will fail silently in 8.x+
Set OPApp = CreateObject("opp.application")
Set proj = OPApp.FileOpen("C:\project.opp")

' ✅ CORRECT - Must call Login first
Set OPApp = CreateObject("opp.application")
OPApp.Login  ' <-- REQUIRED for 8.x+
Set proj = OPApp.FileOpen("C:\project.opp")
```

#### ❌ NOT Calling .TimeAnalyze()
**Impact:** Early/late dates don't update after changes, schedule appears stale
**Fix:** Call `proj.TimeAnalyze` after modifying activities or relationships

```vb
' ❌ WRONG - Dates won't recalculate
act.Duration = "20D"
Debug.Print act.EarlyFinish  ' <-- Shows OLD date

' ✅ CORRECT - Recalculate after changes
act.Duration = "20D"
proj.TimeAnalyze
Debug.Print act.EarlyFinish  ' <-- Shows NEW date
```

#### ❌ NOT Calling .Save()
**Impact:** All changes lost when project closes
**Fix:** Explicitly call Save() or SaveAs() before Shut()

```vb
' ❌ WRONG - Changes will be lost
act.PercentComplete = 100
proj.Shut

' ✅ CORRECT - Save before closing
act.PercentComplete = 100
proj.Save
proj.Shut
```

---

### Import/Export Scripts (Transfer.dat)

#### ❌ Using Lowercase Commands
**Impact:** Scripts fail silently, no error message, no data imported/exported
**Fix:** ALL commands MUST be UPPERCASE

```
❌ WRONG - Will fail silently:
export csv MyExport
table act
field act_id

✅ CORRECT - UPPERCASE required:
EXPORT csv MyExport
TABLE ACT
FIELD ACT_ID
```

#### ❌ Wrong Date Format
**Impact:** Import fails or imports incorrect dates
**Fix:** Use TWO-DIGIT month and day formats (02/05/04 not 2/5/04)

```
❌ WRONG:
DATE_FORMAT %M/%D/%Y  ' May produce single digits

✅ CORRECT:
DATE_FORMAT %M/%D/%Y  ' Always two digits when using %M and %D
Use "02/05/04" in data files, not "2/5/04"
```

#### ❌ Missing LINK Keyword
**Impact:** Related tables don't connect, relationships broken
**Fix:** Use LINK to connect child tables to parent

```
❌ WRONG - Relationships won't import:
TABLE ACT
FIELD ACT_ID

TABLE REL
FIELD PRED_ACT_UID

✅ CORRECT - Link tables properly:
TABLE ACT
FIELD ACT_ID

TABLE REL
LINK SUCC_ID  ' <-- Links relationships to activities
FIELD PRED_ACT_UID
```

---

### Calculated Fields

#### ❌ Using Display Names Instead of Field Names
**Impact:** Calculated field fails with "Unknown field" error
**Fix:** Use database field names (ESDATE) not display names ("Early Start")

```
❌ WRONG:
"Early Start" + |5d|

✅ CORRECT:
ESDATE + |5d|
```

#### ❌ Circular References
**Impact:** Calculated field error or infinite loop
**Fix:** Never reference the calculated field within itself

```
❌ WRONG:
CalcField1 = CalcField1 + 10

✅ CORRECT:
CalcField1 = ORIG_DUR + 10
```

#### ❌ Wrong Constant Syntax
**Impact:** Syntax error or wrong data type
**Fix:** Use proper delimiters for each constant type

```
❌ WRONG:
ESDATE + 5d       ' Missing pipes
ESDATE + "5d"     ' String, not duration
ESDATE + 5        ' Number, not duration

✅ CORRECT:
ESDATE + |5d|     ' Duration properly delimited
```

---

## ✅ REQUIRED PATTERNS

### VBA: Standard Initialization Sequence

```vb
' 1. Create application object
Set OPApp = CreateObject("opp.application")

' 2. Login (REQUIRED for 8.x+)
OPApp.Login

' 3. Configure application (optional)
OPApp.SilentMode = True
OPApp.Show

' 4. Open project
Set proj = OPApp.FileOpen("C:\path\project.opp")

' 5. Perform operations
' ... your code here ...

' 6. Save changes
proj.Save

' 7. Close project
proj.Shut

' 8. Clean up
Set proj = Nothing
Set OPApp = Nothing
```

### VBA: Activity Modification Sequence

```vb
' 1. Disable auto-analyze if making bulk changes
proj.AutoAnalyze = False

' 2. Make changes
Set acts = proj.Activities
For i = 1 To acts.Count
    Set act = acts.Item(i)
    act.Duration = "15D"
    act.PercentComplete = 50
Next i

' 3. Re-enable auto-analyze
proj.AutoAnalyze = True

' 4. Perform time analysis
proj.TimeAnalyze

' 5. Recalculate costs if needed
proj.CalculateCost

' 6. Save
proj.Save
```

### Import/Export: Basic Export Script Pattern

```
EXPORT csv ExportName
REM Description of what this exports
DELIMITED ,
DATE_FORMAT %M/%D/%Y

TABLE ACT
SORT ACT_ID
FIELD ACT_ID
FIELD DESCRIPTION
FIELD ORIG_DUR
FIELD ESDATE
FIELD EFDATE
```

### Import/Export: Basic Import Script Pattern

```
IMPORT csv ImportName
REM Description of what this imports
UPDATE  ' Update existing records
ADD_MISSING_KEYS  ' Create missing referenced records
DELIMITED ,
DATE_FORMAT %M/%D/%Y

TABLE ACT
FIELD ACT_ID
FIELD DESCRIPTION
FIELD ORIG_DUR
```

### Import/Export: XML Export with Relationships

```
EXPORT XML XMLExport
LITERAL_HEADER
<?xml version="1.0" ?>
LITERAL_END

TABLE PRJ
XML_TAG PROJECT

TABLE ACT
XML_TAG ACTIVITIES
XML_TAG ACTIVITY
LINK PRJ  ' Link to project (root)
FIELD ID ID
FIELD ACT_DESC DESCRIPTION
FIELD ESDATE ESDATE

TABLE REL
XML_TAG PREDECESSORS
XML_TAG PREDECESSOR
LINK SUCC_ID  ' Link to activities
FIELD PRED_ACT_UID PRED_ACT_UID
FIELD REL_TYPE REL_TYPE
```

---

## 🚫 ANTI-PATTERNS (Don't Do This)

### VBA: Skipping Error Handling

```vb
' ❌ DANGEROUS - No error handling
Set proj = OPApp.FileOpen("C:\MayNotExist.opp")
proj.Save  ' Crashes if FileOpen failed

' ✅ SAFE - With error handling
On Error Resume Next
Set proj = OPApp.FileOpen("C:\MayNotExist.opp")
If Err.Number <> 0 Then
    MsgBox "Error opening file: " & Err.Description
    Exit Sub
End If
On Error GoTo 0
proj.Save
```

### VBA: Nested Loops Without Performance Optimization

```vb
' ❌ SLOW - No filter, iterates all activities multiple times
Set acts = proj.Activities
For i = 1 To acts.Count
    Set act = acts.Item(i)
    ' ... operations ...
Next i

' ✅ FAST - Apply filter first
Set acts = proj.Activities
acts.Filter = "CRITICAL > 0 AND PCT_COMP < 100"
For i = 1 To acts.Count
    Set act = acts.Item(i)
    ' ... operations ...
Next i
acts.Filter = ""  ' Clear when done
```

### Import/Export: Forgetting UPDATE Mode

```
❌ WRONG - Overwrites all data:
IMPORT csv UpdateData
TABLE ACT
FIELD ACT_ID
FIELD PCT_COMP

' Deletes all activities not in import file!

✅ CORRECT - Updates existing only:
IMPORT csv UpdateData
UPDATE  ' <-- Updates existing, adds new
TABLE ACT
FIELD ACT_ID
FIELD PCT_COMP
```

### Calculated Fields: Complex Expressions Without Variables

```
❌ HARD TO READ:
IIF(DATEDIFFERENCE(ASDATE, TIMENOW()) <= |2d|,
    IIF(DATEDIFFERENCE(ASDATE, TIMENOW()) <= |1d|, "Critical", "Warning"),
    "OK")

✅ READABLE - Use variables:
BEGIN VARIABLES
DaysUntilStart = DATEDIFFERENCE(ASDATE, TIMENOW())
END VARIABLES

IIF(DaysUntilStart <= |2d|,
    IIF(DaysUntilStart <= |1d|, "Critical", "Warning"),
    "OK")
```

---

## 🐛 QUICK TROUBLESHOOTING DECISION TREE

### "Permission denied" or "Access denied" errors
1. ✅ Check: Called `.Login()` after CreateObject?
2. ✅ Check: User has permissions in Deltek EPM Security Administrator?
3. ✅ Check: Menu items enabled (even if invisible)?
4. ✅ Check: File paths accessible and not read-only?

### "Object not found" errors
1. ✅ Check: File paths correct and absolute (not relative)?
2. ✅ Check: Activity/Resource IDs exist in project?
3. ✅ Check: Project/file opened before accessing objects?
4. ✅ Check: Correct case sensitivity for IDs?

### Changes not saving
1. ✅ Check: Called `.Save()` or `.SaveAs()` before `.Shut()`?
2. ✅ Check: File not opened read-only?
3. ✅ Check: File not locked by another user (shared mode)?
4. ✅ Check: Disk space available?

### Schedule not calculating / Dates not updating
1. ✅ Check: Called `.TimeAnalyze()` after modifications?
2. ✅ Check: Calendar file assigned to project?
3. ✅ Check: Activities have calendars assigned?
4. ✅ Check: No circular relationship loops?
5. ✅ Check: StatusDate (Time Now) set correctly?

### Import/Export script not working
1. ✅ Check: Commands in UPPERCASE?
2. ✅ Check: Date format uses two-digit month/day?
3. ✅ Check: Proper LINK keywords for related tables?
4. ✅ Check: Field names match database field names?
5. ✅ Check: File paths accessible?
6. ✅ Check: Delimiter matches data file format?

### Calculated field errors
1. ✅ Check: Using field names (ESDATE) not display names ("Early Start")?
2. ✅ Check: Proper constant syntax (|5d| for durations, {01/01/24} for dates)?
3. ✅ Check: No circular references?
4. ✅ Check: Linked table fields use correct syntax (C1.DESCRIPTION)?
5. ✅ Check: Function names spelled correctly and proper parameters?

### Filters not working
1. ✅ Check: Field names match database names (ACT_ID not "Activity ID")?
2. ✅ Check: Dates wrapped in # delimiters (#1/1/2024#)?
3. ✅ Check: Durations wrapped in pipes (|5d|)?
4. ✅ Check: Text wrapped in quotes ("Complete")?
5. ✅ Check: Proper Boolean operators (AND, OR, NOT)?

---

## 📋 PRE-FLIGHT CHECKLISTS

### Before Running VBA Script

- [ ] `OPApp.Login()` called after CreateObject?
- [ ] Error handling implemented (On Error Resume Next)?
- [ ] `.TimeAnalyze()` called after activity modifications?
- [ ] `.CalculateCost()` called after cost-related changes?
- [ ] `.Save()` called before `.Shut()`?
- [ ] Object variables set to Nothing for cleanup?
- [ ] File paths absolute, not relative?
- [ ] Tested on sample project first?

### Before Running Import/Export Script

- [ ] All commands in UPPERCASE?
- [ ] Date format uses two-digit month/day?
- [ ] LINK keywords used for related tables?
- [ ] UPDATE or UPDATE_ONLY mode specified for updates?
- [ ] Field names match database field names exactly?
- [ ] Test data file format matches script expectations?
- [ ] Backup created before import?
- [ ] Tested on sample project/data first?

### Before Deploying Calculated Field

- [ ] Using field names (ESDATE) not display names?
- [ ] No circular references?
- [ ] Proper constant syntax for all data types?
- [ ] Variables defined in BEGIN/END VARIABLES block?
- [ ] Tested on sample activities first?
- [ ] Result type matches expected output?
- [ ] No references to undefined fields?

---

## 🔑 COMMON ERROR MESSAGES & SOLUTIONS

### "The application object has not been initialized"
**Cause:** Didn't call .Login() in 8.x+
**Solution:** Add `OPApp.Login` immediately after CreateObject

### "Run-time error '13': Type mismatch"
**Cause:** Wrong data type assignment (e.g., string to date field)
**Solution:** Check data types, use proper format (dates need #, durations need |)

### "Run-time error '91': Object variable or With block variable not set"
**Cause:** Trying to use object that wasn't successfully created
**Solution:** Add error checking after Set statements

### "Invalid filter expression"
**Cause:** Wrong field names or syntax in filter
**Solution:** Use database field names, proper delimiters for dates/text

### "Import failed" (Transfer.dat)
**Cause:** Usually lowercase commands or wrong date format
**Solution:** Verify UPPERCASE commands, two-digit dates

### "Unknown field name"
**Cause:** Using display name instead of database field name
**Solution:** Use field name (ESDATE) not display ("Early Start")

---

## 💡 PERFORMANCE TIPS

1. **Disable AutoAnalyze during bulk operations**
   ```vb
   proj.AutoAnalyze = False
   ' ... make many changes ...
   proj.AutoAnalyze = True
   proj.TimeAnalyze  ' Analyze once at end
   ```

2. **Use filters before iterating**
   ```vb
   acts.Filter = "CRITICAL > 0"  ' Reduces iteration count
   For i = 1 To acts.Count
       ' ... process only critical activities ...
   Next i
   ```

3. **Use GetFields() instead of multiple GetField()**
   ```vb
   ' ❌ SLOW - Multiple calls
   id = act.GetField("ACT_ID")
   desc = act.GetField("DESCRIPTN")
   dur = act.GetField("ORIG_DUR")

   ' ✅ FAST - Single call
   Dim fields(2) As String
   fields(0) = "ACT_ID"
   fields(1) = "DESCRIPTN"
   fields(2) = "ORIG_DUR"
   values = act.GetFields(fields)
   ```

4. **Minimize RefreshData usage in shared mode**
   ```vb
   ' Only refresh when absolutely needed
   act.RefreshData = True  ' Forces database read
   value = act.PercentComplete
   ```

---

**Last Updated:** 2025-10-28
**For complete API reference, see:** VBA-API-Reference.md, Import-Export-Reference.md, Calculated-Fields-Reference.md
