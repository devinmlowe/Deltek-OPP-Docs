# Calculated Fields Reference for Deltek Open Plan®

**Purpose:** Complete reference for creating user-defined calculated fields, formulas, and custom expressions.

**⚠️ CRITICAL:** Before using this reference, check [Critical-Warnings-and-Patterns.md](Critical-Warnings-and-Patterns.md) for fatal mistakes like using display names instead of field names.

---

## Overview

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

---

## Elements of Calculated Field Expressions

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

---

## Constants

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

🐛 **Common Issue:** Wrong constant syntax causes syntax errors

✅ **Best Practice:** Always use proper delimiters for each data type

---

## Field Names

**⚠️ IMPORTANT:** Use field names, NOT descriptive column headings.
- Correct: `ESDATE` (not "Early Start")
- Correct: `DESCRIPTN` (not "Description")

**Linked Tables:** Use linking field name + period + field name:
- `C1.DESCRIPTION` — Description for code in C1 field
- `RES_ID.DESCRIPTION` — Description for resource ID
- Linking fields shown with double chevron (») in Fields dialog

**⚠️ WARNING:** Calculated field cannot reference itself (circular reference).

🐛 **Common Issue:** Using "Early Start" instead of ESDATE causes "Unknown field" error

✅ **Best Practice:** Use Fields dialog to verify exact field names

---

## Operators

### Mathematical

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

### Character

- Concatenate: `+`
- Is contained in: `$` (returns logical result)

### Duration Operations

| Operation | Result Type |
|-----------|-------------|
| Duration + Duration | Duration |
| Duration + Date | Date |
| Duration - Duration | Duration |
| Date - Duration | Date |
| Duration / Duration | Decimal |
| Duration / Number | Duration |
| Duration * Number | Duration |

### Relational

- Equal to: `=`
- Not equal to: `<>`
- Greater than: `>`
- Greater than or equal to: `>=`
- Less than: `<`
- Less than or equal to: `<=`
- Contained in: `$`

### Logical

- `AND`
- `OR`
- `NOT`
- `AND NOT`
- Grouping: `()`

---

## User-Defined Variables

Open Plan optimizes expressions by identifying repeated subexpressions. For explicit control and readability, use variable blocks.

**Rules:**
- Variables defined in BEGIN VARIABLES/END VARIABLES block
- Each definition on own line
- Result type determined automatically
- Variable names: no spaces, avoid field/function name conflicts
- Not case sensitive
- Format: `<variable_name> = <expression>`
- **⚠️ CRITICAL:** Variables MUST be defined before referenced (order matters!)
- Can reference previously defined variables

**Simple Example:**
```
BEGIN VARIABLES
X = DATEDIFFERENCE(ASDATE, TIMENOW())
Y = IIF(x<=|2d|, IIF(x<=|1d|, 2, 1), -1)
END VARIABLES

IIF(Y>0, "OK", "Warning")
```

**Complex Enterprise Example (Boeing):**
```
BEGIN VARIABLES
TNplus30d = DATEFORMAT(DATEADD(TIMENOW(), |30d|, CLH_ID), "%M/%D/%Y")
TNplus60d = DATEFORMAT(DATEADD(TIMENOW(), |60d|, CLH_ID), "%M/%D/%Y")
NotSumOrExtSub = IIF(ACT_TYPE<>[Subproject] AND ACT_TYPE<>[External Subproject], "True", "False")
END VARIABLES

IIF(COMPSTAT=[COMPLETE], "Complete",
  IIF(BSDATE={}, "No Baseline",
    IIF(NotSumOrExtSub="True" AND (BSDATE<>{} AND TIMENOW()>=BSDATE), "Late",
      IIF(ESDATE<={} OR TIMENOW()<ESDATE, "Not Started",
        IIF(TIMENOW()>=ESDATE AND TIMENOW()<EFDATE, "On Schedule",
          IIF(TIMENOW()>=EFDATE, "Should Be Complete", "Unknown"))))))
```

**Benefits:**
1. More readable and maintainable
2. Faster parsing of complex expressions
3. Subexpression evaluated once per cell
4. Simplifies debugging - test variables individually

**⚠️ Known Bug/Defect:**
The `<` operator does not parse correctly when used with user-defined variables. Use `>=` instead.

**Example:**
```
❌ WRONG: IIF(MyVariable < |5d|, "Short", "Long")  # Parser error
✅ CORRECT: IIF(MyVariable >= |5d|, "Long", "Short")  # Works properly
```

✅ **Best Practice:** Use variables for complex nested expressions and expressions longer than 500 characters

---

## Calculated Field Functions

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

---

## Function Reference (Detailed)

### ABS()

**Purpose:** Returns absolute value of numeric variable

**Data Type:** Decimal or integer

**Syntax:** `ABS(<value>)`

**Example:** `ABS(-4)` returns `4`

---

### BASELINE_FIELD()

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

### CDOW()

**Purpose:** Returns full day of week

**Data Type:** Character

**Syntax:** `CDOW(<date>)`

**Example:** `CDOW({07OCT04})` returns "Tuesday"

---

### CMONTH()

**Purpose:** Returns full month name

**Data Type:** Character

**Syntax:** `CMONTH(<date>)`

**Example:** `CMONTH({07OCT04})` returns "October"

---

### CTOD()

**Purpose:** Converts character string to DATE data type

**Data Type:** Date

**Syntax:** `CTOD(<String Expression>)`

**Example:** `CTOD(STR(USER_NUM01) + "/" + STR(USER_NUM02) + "/" + STR(YEAR(TIMENOW())))`

If USER_NUM01=12, USER_NUM02=31, Time Now=1/1/2006, returns {12/31/2006}

---

### DATE()

**Purpose:** Returns current date

**Data Type:** Date

**Syntax:** `DATE()`

**Example:** `DATE()` returns current date (e.g., 04OCT04)

---

### DATEADD()

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

### DATEDIFFERENCE()

**Purpose:** Returns difference between two dates as duration

**Data Type:** Duration

**Syntax:** `DATEDIFFERENCE(<date1>, <date2>, <calendar>)`

**Parameters:**
- date1, date2: Date fields or constants
- calendar: Optional (uses project calendar if omitted)

**Example:** `DATEDIFFERENCE({07OCT04}, {08OCT04}, "CAL1")` returns `2d`

---

### DATEFORMAT()

**Purpose:** Returns date in specified format

**Data Type:** Character

**Syntax:** `DATEFORMAT(<date field>, <format string>)`

**Example:** `DATEFORMAT(ESDATE, "%D%A%Y")` returns "07OCT04"

**NOTE:** Uses same format strings as Date Scale Preferences dialog

---

### DAY()

**Purpose:** Returns day of month (1-31)

**Data Type:** Integer

**Syntax:** `DAY(<date>)`

**Example:** `DAY({04OCT04})` returns `4`

---

### DOW()

**Purpose:** Returns day of week (1-7), based on .ini file setting

**Data Type:** Integer

**Syntax:** `DOW(<date>)`

**Example:** `DOW({07OCT04})` returns `3` (Tuesday)

---

### DURATION()

**Purpose:** Returns number of minutes for given duration

**Data Type:** Integer

**Syntax:** `DURATION(<Duration Value>)`

**Example:** `DURATION(|2d|)` returns `960` (8 hours × 2 days × 60 min/hour)

---

### EVAL()

**Purpose:** Evaluates string expression as dynamic calculated field

**Data Type:** Character

**Syntax:** `EVAL(<exp>)`

**Example:** If USER_CHR01="ESDATE", USER_CHR02="EFDATE", ESDATE={04OCT04}, EFDATE={06OCT04}:
```
EVAL('CTOD(USER_CHR02) - CTOD(USER_CHR01)')
```
Returns duration `2d`

---

### FAIL_EVALUATE()

**Purpose:** Tests if expression would fail/return blank

**Data Type:** Boolean

**Syntax:** `FAIL_EVALUATE(<String Expression>)`

**Example:** `FAIL_EVALUATE("C25")` returns TRUE if no code file at index 25

**Note:** Different from "C25 IS_EMPTY" which also returns TRUE if C25 exists but is blank

---

### FISCALPERIOD()

**Purpose:** Returns label for fiscal reporting period containing date

**Data Type:** Character

**Syntax:** `FISCALPERIOD(<Date>, <Reporting Calendar>)`

**Parameters:**
- Date: Search date
- Reporting Calendar: Optional (defaults to project reporting calendar)

**Example:** `FISCALPERIOD(ESDATE)` or `FISCALPERIOD(ESDATE, "REPORTING_CALENDAR_NAME")`

Returns LABEL from reporting calendar containing supplied date

---

### FORMAT_HEADING_ITEM()

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

### GET_ASSGNS()

**Purpose:** Returns assignment data for activity

**Data Type:** Character

**Syntax:** `GET_ASSGNS(<fieldname1>[|<fieldname2>...])`

**Format:** Fields separated by commas, records by semicolons

**Example:** `GET_ASSGNS("RES_ID|RES_LEVEL")` returns "ENG,1.00;TECH.MARY,2.00"

**NOTE:** Remove table portion of fieldname when using this function

---

### GET_CHILDREN()

**Purpose:** Returns list of fields from immediate children of hierarchical record

**Data Type:** Character

**Syntax:** `GET_CHILDREN(<FieldList>)`

**Format:** Pipe-delimited field list, comma-separated values, semicolon-separated records

**Example:** `GET_CHILDREN("ACT_ID|DESCRIPTION")` for activity "1.01" returns:
```
"1.01.01,First child of 1.01;1.01.02,Second child of 1.01"
```

---

### GET_COSTS()

**Purpose:** Returns cost records for activity

**Data Type:** Character

**Syntax:** `GET_COSTS(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_COSTS("ACWP_CST|ACWP_QTY")` returns "1000.00,1.00;4000.00,2.00"

---

### GET_FIELD()

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

### GET_FIRST_RECORD_IN_SUMMARY()

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

### GET_NOTE()

**Purpose:** Returns activity, resource, or code note

**Data Type:** Character

**Syntax:** `GET_NOTE(<category>)`

**Parameters:**
- category: Note category name (optional, defaults to default category)

**Examples:**
- `GET_NOTE("Document")` returns Document category note
- `GET_NOTE("")` returns default category note

---

### GET_PREDS()

**Purpose:** Returns predecessor data for activity

**Data Type:** Character

**Syntax:** `GET_PREDS(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_PREDS("PRED_ACT_ID|REL_TYPE")` returns "A100,Finish to Start"

---

### GET_RELATED()

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

### GET_RELATED_COUNT()

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

### GET_RISKS()

**Purpose:** Returns risk data for activity

**Data Type:** Character

**Syntax:** `GET_RISKS(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_RISKS("ESDATE1")` returns each early start date from risk analysis

---

### GET_SUCCS()

**Purpose:** Returns successor data for activity

**Data Type:** Character

**Syntax:** `GET_SUCCS(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_SUCCS("SUCC_ACT_ID|REL_TYPE")` returns "A200,Finish to Start"

---

### GET_USAGES()

**Purpose:** Returns usage records for activity

**Data Type:** Character

**Syntax:** `GET_USAGES(<fieldname1>[|<fieldname2>...])`

**Example:** `GET_USAGES("RES_ID|RES_USED")` returns "ENG,24"

---

### GO_MONTH()

**Purpose:** Returns date N months before/after specified date

**Data Type:** Date

**Syntax:** `GO_MONTH(<date>, <integer>)`

**Example:** `GO_MONTH({04OCT04}, -2)` returns {04AUG04}

---

### HAS_NOTE()

**Purpose:** Returns logical indicating if note attached

**Data Type:** Logical

**Syntax:** `HAS_NOTE(<string>)`

**Parameters:**
- string: Optional category name

**Examples:**
- `HAS_NOTE()` returns TRUE if any note attached
- `HAS_NOTE("Scope")` returns TRUE if note in "Scope" category attached

---

### IIF()

**Purpose:** Conditional processing (if-then-else)

**Data Type:** Any (matches return value type)

**Syntax:** `IIF(<logicexp>, <iftrue>, <iffalse>)`

**Operation:** If logicexp is TRUE, return iftrue; else return iffalse

**Example:** `IIF(ESDATE > {01JUL01}, "Underway", "Planned")` returns "Planned" if ESDATE is {19JUN01}

✅ **Best Practice:** Nest IIF() functions for multi-condition logic

---

### INLIST()

**Purpose:** Returns logical indicating if value in list

**Data Type:** Logical

**Syntax:** `INLIST(<search>, <value1>, <value2>, ...)`

**Examples:**
- `INLIST(MONTH(ESDATE), 1, 4, 7, 10)` returns FALSE if month is June
- `INLIST("CHRIS", GET_ASSGNS("LOCAL(RES_ID)"))` returns TRUE for activities where CHRIS assigned

**NOTE:** Character argument interpreted as comma-delimited list

---

### INSTR()

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

### LEFT()

**Purpose:** Returns leftmost N characters

**Data Type:** Character

**Syntax:** `LEFT(<string>, <int>)`

**Example:** `LEFT("SITE COORDINATION AND DESIGN", 4)` returns "SITE"

---

### LEN()

**Purpose:** Returns length of variable or constant

**Data Type:** Integer

**Syntax:** `LEN(<data>)`

**Example:**
- `LEN("Dig Hole")` returns `8`
- `LEN(DESCRIPTION)` where DESCRIPTION="EXCAVATION" returns `10` (not field width)

---

### LEVEL()

**Purpose:** Returns hierarchical level of ID or code

**Data Type:** Integer

**Syntax:** `LEVEL(<ID>)`

**Example:** `LEVEL(C1)` returns level of C1 in code structure

---

### LOCAL()

**Purpose:** Returns local portion of ID/code (not shared with siblings)

**Data Type:** Character

**Syntax:** `LOCAL(<ID>, <level>)`

**Parameters:**
- ID: Character variable or constant
- level: Optional - specific level for local portion

**Example:** `LOCAL(C1)` returns rightmost portion of C1

---

### LOWER()

**Purpose:** Converts uppercase to lowercase

**Data Type:** Character

**Syntax:** `LOWER(<string>)`

**Example:** `LOWER("Dig Hole")` returns "dig hole"

---

### LTRIM()

**Purpose:** Trims leading spaces

**Data Type:** Character

**Syntax:** `LTRIM(<string>)`

---

## Additional Functions

For complete reference on additional functions including MAX, MIN, MID, MONTH, NEWLINE, NUMBER_FORMAT, OCCURS, PARENT, RECORD_NUMBER, RIGHT, ROUND, SPACE, SQRT, STR, STRTRAN, STUFF, SUBSTR, TIMENOW, TRIM, UPPER, USER_ID, VAL, and YEAR, refer to the full DeltekOpenPlanDeveloperGuide.md or Open Plan documentation.

---

## 🏢 Boeing Enterprise Best Practices

### Naming Convention
**Format:** `YourID_CF_Description`

**Examples:**
- `DLOWE_CF_StatusFlag`
- `JSMITH_CF_ScheduleVariance`
- `ABROWN_CF_ResourceCount`

**Rules:**
- Maximum 60 characters for calculated field name
- Use your BEMS ID or last name as prefix
- Use underscore separators
- Keep description clear and concise
- Avoid special characters beyond underscore

### Expression Length Limits
- **Maximum:** 4,096 characters per calculated field expression
- **Recommended:** Under 2,000 characters for maintainability
- **Strategy:** Break very complex logic into multiple calculated fields

### Development Workflow
1. **Build Bottom-Up:** Start with inner functions, test, then add outer logic
2. **Test Incrementally:** Test each piece before combining
3. **Use Variables:** For expressions over 500 characters, use BEGIN VARIABLES block
4. **Document:** Add comments in description field explaining complex logic
5. **Peer Review:** Have another scheduler review complex formulas

### Text Editor Requirements
**⚠️ CRITICAL:** Use Notepad or similar plain text editor. DO NOT use Microsoft Word!

**Why?**
- Word formats quotation marks to "smart quotes" (ASCII characters: " " instead of ")
- These ASCII characters cause parser errors that are difficult to diagnose
- Word may add hidden formatting characters

**Recommended Editors:**
- Windows Notepad
- Notepad++
- Visual Studio Code
- Any plain text editor without auto-formatting

**Safe Workflow:**
1. Write complex expressions in plain text editor
2. Test pieces in Open Plan calculated field dialog
3. Copy/paste from plain text editor into Open Plan
4. Never copy from Word, PowerPoint, or formatted documents

---

## 🐛 Debugging Guide

### Most Common Errors (Boeing Training Data)

#### Error #1: "The parser was not able to interpret the calculated field expression"
**Cause:** Syntax error in expression

**Common Mistakes:**
- Missing or mismatched parentheses
- Wrong constant delimiters: `{date}`, `|duration|`, `"text"`, `[enum]`
- Misspelled function names
- Using smart quotes from Word instead of plain quotes
- Missing operators between terms

**Fix:**
1. Check all parentheses are balanced
2. Verify constant syntax: `{01JAN01}`, `|5d|`, `"text"`, `[ASAP]`
3. Use plain text editor (Notepad), not Word
4. Break complex expression into smaller parts and test each

**Example:**
```
❌ WRONG: IIF(ESDATE > 01JAN01, "Late", "OK")  # Missing { }
✅ CORRECT: IIF(ESDATE > {01JAN01}, "Late", "OK")

❌ WRONG: IIF(DURATION > 5d, "Long", "Short")  # Missing | |
✅ CORRECT: IIF(DURATION > |5d|, "Long", "Short")
```

#### Error #2: "Incompatible data types in expression"
**Cause:** Type mismatch in operation or missing enumerated value brackets

**Common Mistakes:**
- Comparing different data types (date to string, duration to number)
- Missing square brackets on enumerated values
- IIF() branches returning different data types
- Wrong function parameter types

**Fix:**
1. Check enumerated values have square brackets: `[ASAP]`, `[Start Milestone]`
2. Ensure IIF() true/false branches return same data type
3. Verify function parameters match expected types
4. Use type conversion functions: CTOD(), STR(), VAL()

**Examples:**
```
❌ WRONG: IIF(ACT_TYPE = ASAP, "OK", "Not OK")  # Missing [ ]
✅ CORRECT: IIF(ACT_TYPE = [ASAP], "OK", "Not OK")

❌ WRONG: IIF(ESDATE > {01JAN01}, "Late", 0)  # Mixed types
✅ CORRECT: IIF(ESDATE > {01JAN01}, "Late", "On Time")
```

### "Unknown field name" Error
1. ✅ Check: Using field name (ESDATE) not display name ("Early Start")?
2. ✅ Check: Field exists in current table?
3. ✅ Check: Linked table syntax correct (C1.DESCRIPTION)?
4. ✅ Check: Use Fields dialog to verify exact field name

### "Circular reference" Error
1. ✅ Check: Calculated field not referencing itself?
2. ✅ Check: No chain of calculated fields that loops back?

### User-Defined Variable Errors
1. ✅ Check: Variable defined before referenced?
2. ✅ Check: Variable name doesn't conflict with field/function name?
3. ✅ Check: Using `>=` instead of `<` operator (known bug)?
4. ✅ Check: BEGIN VARIABLES and END VARIABLES properly spelled?

### Wrong Result/Blank Cell
1. ✅ Check: Function returns expected data type?
2. ✅ Check: All IIF() branches return same data type?
3. ✅ Check: Empty/null values handled properly?
4. ✅ Check: Calendar specified for date operations?
5. ✅ Check: Field has data in current record?

### Performance Issues (Expression Recalculating Slowly)
1. ✅ Check: Using variables for repeated subexpressions?
2. ✅ Check: Avoiding nested GET_* functions?
3. ✅ Check: Expression length under 2,000 characters?
4. ✅ Check: Not using EVAL() unnecessarily (slow function)?

---

## ✅ Best Practices Summary

1. **Use field names, not display names** - ESDATE not "Early Start"
2. **Use variables for complex expressions** - Makes formulas readable and faster
3. **Avoid circular references** - Field cannot reference itself
4. **Use proper constant syntax** - {date}, |duration|, "text", [enum]
5. **Test on sample data first** - Verify formula logic before deploying
6. **Document complex formulas** - Add comments in REM field
7. **Use IIF() for conditional logic** - Nest carefully for multi-conditions
8. **Leverage GET_* functions** - Access related table data efficiently
9. **Check result types** - Ensure operations produce expected data types
10. **Use Fields dialog** - Verify exact field names and types

---

## Common Formula Patterns

### Schedule Variance
```
DATEDIFFERENCE(BASELINE_FINISH, EFDATE)
```

### Cost Variance
```
ACTUAL_COST - BUDGET_COST
```

### Days Until Start
```
DATEDIFFERENCE(DATE(), ESDATE)
```

### Status Flag
```
IIF(PCT_COMP = 100, "Complete",
  IIF(PCT_COMP > 0, "In Progress", "Not Started"))
```

### Critical and Behind Schedule
```
CRITICAL > 0 AND EFDATE > BASELINE_FINISH
```

### Resource List
```
GET_ASSGNS("RES_ID")
```

### Hierarchical Summary
```
GET_CHILDREN("ACT_ID|PCT_COMP")
```

---

*For critical warnings and troubleshooting, see [Critical-Warnings-and-Patterns.md](Critical-Warnings-and-Patterns.md)*
*For VBA automation, see [VBA-API-Reference.md](VBA-API-Reference.md)*

**Last Updated:** 2025-10-28
