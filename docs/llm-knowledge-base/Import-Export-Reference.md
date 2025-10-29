# Import/Export Reference for Deltek Open Plan®

**Purpose:** Complete Transfer.dat scripting reference for custom import/export operations.

**⚠️ CRITICAL:** Before using this reference, check [Critical-Warnings-and-Patterns.md](Critical-Warnings-and-Patterns.md) for fatal mistakes. **ALL COMMANDS MUST BE UPPERCASE** or scripts fail silently.

---

## Overview

When you use General Import/Export from the Integration tab, Open Plan uses scripts defined in the **Transfer.dat** file (located in the local system folder). This is a text file that can be edited to create custom import/export specifications.

**CRITICAL:** All import/export command scripts are **case-sensitive**. Commands must be entered in UPPERCASE or you may receive errors.

### Transfer.dat File Structure

Transfer.dat can contain:
- Script definitions for import/export operations
- Pointers to separate include files
- Each line must start with a command word

---

## Default Import Scripts

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

---

## Default Export Scripts

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

---

## Basic Import/Export Script Examples

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

---

## Core Commands

### IMPORT Command

Syntax: `IMPORT <extension> <name>`

Defines start of import script. Everything from this command to next IMPORT/EXPORT or EOF defines the script.

- `<extension>`: Default file extension (can use wildcard `*`)
- `<name>`: Appears in Open Plan Import dialog

Examples:
```
IMPORT csv Excel file
IMPORT * Excel
```

🐛 **Common Issue:** Lowercase "import" causes silent failure

---

### EXPORT Command

Syntax: `EXPORT <extension> <name>`

Defines start of export script. Everything from this command to next IMPORT/EXPORT or EOF defines the script.

- `<extension>`: Default file extension (can use wildcard `*`)
- `<name>`: Appears in Open Plan Export dialog

Examples:
```
EXPORT XML Extended Markup Language (HTML)
EXPORT * XML
```

🐛 **Common Issue:** Lowercase "export" causes silent failure

---

### TABLE Command

Syntax: `TABLE <optabletype>`

Indicates the current table for the script. Everything from this to next TABLE, IMPORT, EXPORT, or EOF applies to this table.

Table types: ACT, REL, ASG, CST, USE, RES, AVL, CLH, CLR, SCA, CNN, PRJ, etc.

Example:
```
TABLE ACT
```

---

### RECORD_TYPE Command

Syntax: `RECORD_TYPE <recordtype>`

Defines identifier for record type, assumed to be first item of each record.

---

### FIELD Command

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

---

### INCLUDE Command

Syntax: `INCLUDE <filename>`

Allows script files to be included by reference. Include files may be nested but must NOT contain IMPORT or EXPORT commands.

Example:
```
INCLUDE ACTDATA1.XFR
```

---

## Additional Commands

### ADD_MISSING_KEYS

For imports, automatically create records for missing data. Example: If import has resource assignments but resource file is empty, Open Plan creates the referenced resource records automatically.

```
ADD_MISSING_KEYS
```

✅ **Best Practice:** Use this for imports with related tables to avoid errors

---

### DATE_FORMAT

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

**⚠️ IMPORTANT:** Open Plan requires **two-digit** month, day, and time formats when importing from comma-delimited files. Example: Use `02/05/04` not `2/5/04`.

🐛 **Common Issue:** Single-digit dates (2/5/04) cause import failures

---

### DELIMITED

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

---

### DURATION_FORMAT

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

---

### FIELD_SPECIAL

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

---

### FILTER

Syntax: `FILTER [<filtername>] [<filterdefinition>]`

Applies filter to export data. Filter definition should contain no spaces.

Example:
```
FILTER CRITICAL>0ANDPCT_COMP<100
```

---

### FIXED

Syntax:
```
FIXED <fieldwidth>
FIXED <margin> <fieldwidth>
```

Defines file as fixed-format.

- One parameter: `<fieldwidth>` = characters for data record type
- Two parameters: `<margin>` = characters to ignore before record type, `<fieldwidth>` = record type width

Must precede first TABLE command. If absent, file assumed comma-delimited.

---

### HEADER

Syntax: `HEADER <headerrecordtype>`

Defines record type of header record. Header records indicate fields to import/export and their order using identifiers matching `<header>` parameter of FIELD command.

---

### LINK

Syntax: `LINK [<fieldname>]`

Links current table to previously defined table.

Example - link assignment table to activity table:
```
LINK ID
```

Special form without field name links to most recently defined table of same type on record-by-record basis.

Tables may be linked to linked tables, creating chains of any length.

🐛 **Common Issue:** Missing LINK causes relationships to not import

✅ **Best Practice:** Always use LINK for related tables (REL to ACT, ASG to ACT, etc.)

---

### LITERAL_END

Syntax: `LITERAL_END`

Terminates literal data introduced by LITERAL_HEADER or LITERAL_FOOTER.

---

### LITERAL_FOOTER

Syntax: `LITERAL_FOOTER`

Introduces lines to transfer at end of output for current table. Terminated by LITERAL_END.

Example:
```
LITERAL_FOOTER
</ACTIVITIES>
</PROJECT>
LITERAL_END
```

---

### LITERAL_HEADER

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

---

### MPX_CALENDAR_DEFINITION

Syntax: `MPX_CALENDAR_DEFINITION`

Processes calendar definitions in format required by .mpx record type 20. Used only with CLH table type.

---

### MPX_CALENDAR_HOURS

Syntax: `MPX_CALENDAR_HOURS`

Processes standard working hours in format required by .mpx record type 25. Used only with CLH table type.

---

### MPX_CALENDAR_EXCEPTIONS

Syntax: `MPX_CALENDAR_EXCEPTIONS`

Processes calendar exceptions (holidays) in format required by .mpx record type 26. Used only with CLH table type.

---

### REM

Syntax: `REM <comment>`

Adds comment lines to script. Open Plan ignores lines starting with REM.

Example:
```
REM This exports critical activities only
```

✅ **Best Practice:** Document complex scripts with REM comments

---

### SKIP

Syntax: `SKIP <numberoffields>|<numberofcharacters>`

Skips items before next defined field.

- Delimited: `<numberoffields>` to skip
- Fixed-format: `<numberofcharacters>` to skip

Typically used to skip unnecessary fields during import.

---

### SORT

Syntax: `SORT [-]<sort_field1>[,[-]<sort_field2>[,...]]`

Applies sort to export data. Negative sign indicates descending sort.

Example:
```
SORT ACT_ID
SORT C1.codedesc,(ssdate-esdate)
```

---

### UPDATE

Syntax: `UPDATE`

For import, specifies existing records matching incoming ones should be updated. Incoming data not matching existing records are added as new.

Apply to each table that needs updating.

🐛 **Common Issue:** Without UPDATE, import deletes existing records not in import file

---

### UPDATE_ONLY

Syntax: `UPDATE_ONLY`

For import, incoming data only updates existing records. Incoming data without match is ignored.

**IMPORTANT:** When tables are parent-child linked (e.g., activities and assignments):
- If parent is in UPDATE or UPDATE_ONLY mode and child is not, all existing children are deleted before adding new child records
- When LINK used on same table type, second instance automatically set to UPDATE_ONLY mode

---

### UPDATE_REMAINING

Syntax: `UPDATE_REMAINING`

For CST table imports only. Updates Remaining Quantity on activity Resource Assignments when Actual Quantities imported for Activity Resources where resource configured for manual progress (Progress Based On Activity Progress disabled).

---

## Table Types

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

---

## XML Import/Export Commands

### ATTRIBUTE Keyword

Synonym for FIELD, supports XML nomenclature to distinguish from ELEMENT.

---

### ELEMENT Keyword

Similar to FIELD but causes field to be transferred as XML element rather than attribute.

Example output:
```xml
<ACTIVITY <ID>1</ID>
DESCRIPTION="Environment Management System"
ESDATE="02Jan97" />
```

**NOTE:** Elements always output first regardless of definition order.

---

### HIERARCHICAL Keyword

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

---

### STYLESHEET Keyword

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

---

### XML_TAG Keyword

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

---

### FIELD Keyword (XML Context)

If XML_TAG specified for table, second parameter on FIELD contains XML attribute name. May be omitted if same as Open Plan field name.

**Export:** Blank character string attributes not exported

**Import:**
- Not all attributes required for every element
- At least one attribute needed or no record created
- Even ID can be missing (uses auto-numbering)
- Attributes in input file but not defined in script are ignored (no log message)
- Relationships cannot be created prior to both activities

---

### LINK Keyword (XML Context)

Embeds table inside parent table element.

Special form:
```
LINK PRJ
```

Embeds elements inside project element (root element).

**NOTE:** In valid XML, all data must appear within root tag. Every table except PRJ must contain LINK keyword linking to PRJ or previously-defined table.

---

### LITERAL_HEADER & LITERAL_FOOTER (XML Context)

Can be defined outside table definition scope. Facilitates XML declaration and heading information.

---

## Import/Export Considerations

1. Log messages may reference lines that don't uniquely identify issue
2. On input, can nest top-level tags within other tags (works but doesn't imply hierarchy)
3. On export, no way to nest activities within other activities
4. Import log messages may not pinpoint error location (CR/LF has no significance in XML)

---

## XML Complete Examples

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

---

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

---

## 🐛 Debugging Guide

### Script Not Running
1. ✅ Check: All commands in UPPERCASE?
2. ✅ Check: Transfer.dat file in correct location (system folder)?
3. ✅ Check: Script name appears in Import/Export dialog?

### Fields Not Importing
1. ✅ Check: Date format uses two-digit month/day?
2. ✅ Check: Field names match database field names exactly?
3. ✅ Check: LINK keywords present for related tables?
4. ✅ Check: Delimiter matches data file?

### Relationships Not Importing
1. ✅ Check: LINK SUCC_ID or LINK PRED_ACT_UID present?
2. ✅ Check: Both activities exist before creating relationship?
3. ✅ Check: UPDATE mode if updating existing relationships?

### XML Import Errors
1. ✅ Check: Root element (PRJ table with XML_TAG PROJECT)?
2. ✅ Check: All tables have LINK to PRJ or parent table?
3. ✅ Check: XML_TAG defined for all tables?
4. ✅ Check: Valid XML structure (balanced tags)?

---

## ✅ Best Practices Summary

1. **Always use UPPERCASE for commands** - This is the #1 cause of failures
2. **Test scripts on sample data first** - Catch errors before production
3. **Use REM for documentation** - Make scripts maintainable
4. **Include header records for clarity** - Makes data files self-documenting
5. **Validate date formats** - Two-digit month/day required (02/05/04)
6. **Check for special characters in data** - Quotes, commas, pipes can cause issues
7. **Use filters to limit scope** - Export only what's needed
8. **Log errors for troubleshooting** - Review import/export logs
9. **Create backups before imports** - Always have rollback option
10. **Use UPDATE mode carefully** - Understand update vs. replace behavior

---

## Important Import/Export Notes

**Problem Characters:** These symbols may cause issues as delimiters in fields:
- Quotation marks (")
- Commas (,)
- Piping symbols (|)
- Semicolons (;)

**Syntax Conventions:**
- `<...>` Required parameters
- `[...]` Optional parameters
- `[...|...]` Choice of optional parameters

---

*For critical warnings and troubleshooting, see [Critical-Warnings-and-Patterns.md](Critical-Warnings-and-Patterns.md)*
*For VBA import/export operations, see [VBA-API-Reference.md](VBA-API-Reference.md)*

**Last Updated:** 2025-10-28
