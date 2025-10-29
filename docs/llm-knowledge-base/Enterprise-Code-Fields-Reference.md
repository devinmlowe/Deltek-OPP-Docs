# Enterprise Code Fields Reference for Deltek Open Plan®

**Purpose:** Boeing enterprise-level code field standards for CSPR/Cobra integration and program compliance.

**Source:** CSPR_OPP_DataDictionary.csv - Boeing IP&S Home Office Tools Team

**⚠️ CRITICAL:** These code field assignments are mandatory or recommended for Boeing programs integrating with Cobra, CSPR, and other enterprise systems.

---

## Overview

Open Plan® Professional provides 99 code fields (C1-C99) at both the Activity and Project levels. Boeing enterprise standards define how these fields should be used across programs to ensure:

- **Cost/Schedule Integration** with Cobra reporting system
- **CSPR System Requirements** for program management
- **Consistency** across Boeing programs
- **Automated Data Exchange** between systems
- **Compliance** with BDS process requirements

**Code Field Maintenance:**
- **IP&S Home Office Tools Team:** Maintains predetermined code files (limited value sets)
- **Program Focal:** Maintains program-specific code files and assignments

---

## Code Field Requirement Levels

| Level | Description |
|-------|-------------|
| **Mandatory** | Required for CSPR 3 system integration |
| **Recommended (Mandatory by Process)** | Required by Boeing processes/procedures |
| **Mandatory by Process** | Required when specific conditions met (e.g., MES interface, EtBB, Level 3 Change Control) |
| **Optional** | Program discretion |

---

## Activity-Level Code Fields (Critical Fields)

### C1 - Work Breakdown Structure (WBS)

**CSPR Requirement:** Recommended (Mandatory by Process)

**Maintained By:** Program Focal

**Baselined:** Yes

**Description:** Work Breakdown Structure Code. Punctuated Significant or Fixed Format in OPP.

**Requirements:**
- Either matches WBS in Cobra or is at a lower level
- For BDS Programs using DFAR with Acumen and/or JSON: **MUST use this code field**
- Field is NOT populated automatically
- During OPP/Cobra mapping, the Cobra assigned WBS is populated in the `C_USERCHR01` field of the Control Account code value

**Examples:** `1.3.1`, `5150`

**References:** IPSG-9000-007 WBS, OBS, and CAM Fields

---

### C2 - Organizational Breakdown Structure (OBS)

**CSPR Requirement:** Recommended (Mandatory by Process)

**Maintained By:** Program Focal

**Baselined:** Yes (Display Only)

**Description:** Responsibility Breakdown Structure. Breaks down responsibility from Program Manager to Assembly & Integration Teams (AIT's) and/or Integrated Product Teams (IPT's) down to Responsible Control/Cost Account Managers (CAM's).

**Requirements:**
- Either matches OBS in Cobra or is at a lower level
- For BDS Programs using DFAR with Acumen and/or JSON: **MUST use this code field**
- Field is NOT populated automatically
- During OPP/Cobra mapping, the Cobra assigned OBS is populated in the `C_USERCHR02` field of the Control Account code value

**Examples:** `Wing IPT Manager`, `Fuselage AIT Manager`, `A00`, `AA0`

**References:** IPSG-9000-007 WBS, OBS, and CAM Fields

---

### C3 - Percent Complete Type

**CSPR Requirement:** Recommended (Mandatory by Process)

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** Identifies which type of Quantifiable Assess Percent Complete (QA%C) the activity is using. Used when backup data will be defined directly in OPP; otherwise leave blank.

**Predetermined Values:** `R`, `S`, `U`, `T`

**References:** IPSG-9000-019 Quantifiable Backup Data

---

### C4 - Control Account ⚠️ MANDATORY

**CSPR Requirement:** Mandatory

**Maintained By:** Program Focal

**Baselined:** Yes

**Description:** **Unique Identifier within a Cobra Reporting Master**, which identifies a control account. A control account is assigned to an intersection of a WBS and OBS and is the lowest level in Cobra where Budget (BCWS), Earned Value (BCWP) and Actuals (ACWP) are integrated. **Required for Cost/Schedule Integration.**

**Requirements:**
- Must match Control Accounts flagged for OPP in Cobra
- **⚠️ CRITICAL:** This table should ONLY be populated using OPP/Cobra mapping. Manual inputs can cause integration problems.
- During OPP/Cobra mapping, the WBS, OBS, CAM, Planner information assigned in Cobra is automatically populated in Code Value User Defined Fields: `C_USERCHR01`, `C_USERCHR02`, `C_CAM`, `C_PLANNER`

**Example:** `HAA10AAL`

**References:** IPSG-9000-007 WBS, OBS, and CAM Fields

---

### C5 - Integrated Master Plan (IMP)

**CSPR Requirement:** Recommended (Mandatory by Process)

**Maintained By:** Program Focal

**Baselined:** Yes

**Description:** Index # and description of Integrated Master Plan/Integrated Master Schedule item. Typically a punctuated significant code but could be fixed format.

**Requirements:**
- Must match approved IMP
- Required for programs with an IMP requirement

**Example:** `1.1.1`

**References:** IPSG-9000-012 IMS Content Requirements

---

### C6 - Cost Class ⚠️ MANDATORY

**CSPR Requirement:** Mandatory

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Baselined:** Yes

**Description:** Cost Class used to document and control baseline changes and approvals. **Mandatory for CD processing & Change Control.**

**Predetermined Values:** `CB`, `C1`, `CK`

---

### C7 - Planner Name

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Baselined:** Yes

**Description:** Used to assign the planner's name to the work package. The description should contain the planner's name (last, first, MI recommended).

**Format:**
- **Historical:** `Last_First>BEMS`
- **NOW:** `Last_First_BEMS`

**Example:** `Smith_John_12345`

---

### C8 - Status Focal

**CSPR Requirement:** Recommended (Mandatory by Process)

**Maintained By:** Program Focal

**Baselined:** Yes

**Description:** Used to assign a Status Focal (CAM or Delegate) to an activity to allow statusing by a Status Focal in CAM Vis. The description should contain the status focal's name (last, first, MI recommended).

**Format:**
- **Historical:** `Last_First>BEMS`
- **NOW:** `Last_First_BEMS`

**Notes:**
- Field is NOT populated automatically
- During OPP/Cobra mapping, the Cobra assigned CAM is populated in the `C_CAM` field of the Control Account code value

**References:** IPSG-9000-007 'WBS, OBS, and CAM Fields' and IPSG-9000-012 'IMS Content Requirements'

---

### C9 - Integrated Scheduler

**CSPR Requirement:** Recommended (Mandatory by Process)

**Maintained By:** Program Focal

**Baselined:** Yes

**Description:** Used to assign the scheduler's name to the work package. The code description should be the scheduler's BEMS ID.

**Format:**
- **Historical:** `Last_First>BEMS`
- **NOW:** `Last_First_BEMS`

**References:** IPSG-9000-003 Integrated Scheduler Code Field

---

### C10 - Program Breakdown Structure

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Description:** An alternate breakdown structure to segregate work within an IPT.

**Examples:** `ARC`, `MB1`

---

### C11 - Statement of Work Paragraph ⚠️ MANDATORY

**CSPR Requirement:** Mandatory

**Maintained By:** Program Focal (Predetermined)

**Baselined:** Yes

**Description:** Index # and description of Program Statement of Work Paragraph. Typically a punctuated significant code but could be fixed format.

**Example:** `1.1.1`

**References:** IPSG-9000-012 IMS Content Requirements, DID, IPMR and IPMDAR

---

## Activity-Level Code Fields (Standard Fields)

### C12 - Contract Line Item Number

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Description:** Identifies Line Item Code in Contract which provides contractual funding for task being performed.

**Example:** `0001AA`, `0002AA`, `0002AB`, `0003`

---

### C13 - Configuration Item or Model Code

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Baselined:** Yes

**Description:** Identifier for hardware or software configuration end item such as Aircraft model code.

**Examples:** `74A10000-1001`, `13C OFP`

---

### C14 - Contract Data Requirements List

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Description:** Identifies Code corresponding to Contract Data Requirement List for activity being performed.

**Examples:** `A001`, `B003`, `M012`

---

### C15-C19 - Work Definition Document (WDD) Fields

**CSPR Requirement:** Optional

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Required For:** MES Interface

#### C15 - WDD Material Type
**Examples:** `Alum`

**Description:** Identifies type of material utilized on longest lead, most complex part on a particular drawing. Used in conjunction with part type, size & complexity code to identify applicable standard tasks and duration times for drawing/build to activities.

#### C16 - WDD Part Type
**Examples:** `Machine Extrusion`

**Description:** Identifies type of manufacturing process utilized on longest lead, most complex part on a particular drawing.

#### C17 - WDD Size & Complexity
**Examples:** `S3`

**Description:** Fields Size and Complexity have been combined to produce a two character code: Small/Medium/Large for first position, Simple...Extremely Complex 1-4. Code example `S3` represents Simple (size) that is Complex.

#### C18 - WDD Type of Change/Key Deliverables
**Examples:** `DCN`, `EO`, `New`, `AV`, `TST`, `Fab`

**Baselined:** Yes

**Description:** Code which identifies type of change. Determines standard template durations for Change Coordination & Release activities. Also used in ReMICS.

#### C19 - WDD Kind
**Examples:** `schematic`, `mock-up`, `FT`

**Description:** Used to categorize & identify major kinds of IPT Documents such as Production Drawings, Schematics, Mock-Up Drawings, Ground test Drawings, Flight Test Drawings, Procurement Specs, etc.

---

### C20 - Make/Buy

**CSPR Requirement:** Mandatory by Process

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** Make/Buy code field to identify whether the part is to be made ('M') or purchased ('B').

**Predetermined Values:** `M`, `B`

**Required When:** Project is Associated with MES

---

### C21 - T&E Reference Number

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Description:** Identifier assigned by T&E to specific test activities. Facilitates filtering and sorting of specific sequences of test activities for reporting and metrics.

---

### C22 - Effectivity

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Description:** Identifies unique code of Ground Test Article for 1st applicable Test Effectivity of particular non-recurring activity.

**Examples:** `FT50`, `IB50`

---

### C23 - Alternate Symbology

**CSPR Requirement:** Optional

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** Used to identify Billing Events and schedule replans. Combined with Graphic Summarization Type and Tiered code files, flags how activity should be depicted graphically when shown on summary bars/symbols. Also used with ReMICS to show Scheduling Metrics and Important Milestones.

**Examples:** `BILLING $ EVENT`, `PROPOSED RESCHEDULE`

---

### C24 - Graphic Physical Position

**CSPR Requirement:** Optional

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** This field allows a user, with an appropriate template, to "force" the milestone symbol, milestone description, and milestone date (text) to appear at various locations above/below a task bar.

**Examples:** `UL1`, `LR2`

---

### C25 - Reporting Activity

**CSPR Requirement:** Optional

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** Identifies the activity to be copied into the Month End OPP Project using the Month-End Closing tool.

**Predetermined Values:** `Yes`

---

### C26 - 60/90 Day and Criticality Flag

**CSPR Requirement:** Optional

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Baselined:** Yes

**Description:** Flags criticality of activity from an overall program perspective (cost/schedule/technical/risk). Goes to ReMICS.

**Examples:** `60/90`, `CP`

---

### C27 - Linked Project Change Code

**CSPR Requirement:** Optional

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** Informs the user as to status of the Virtual Task after it has been added or updated.

**Examples:** `Updated`, `Not Found`, `Summary`

---

### C28 - Tiering capable 60/90 Day and Criticality Flag

**CSPR Requirement:** Optional

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** Identifies activity to be included within a Scheduling Metrics count and the definition of the count.

**Examples:** `60/90E`, `60/90I`, `60/90P`, `CPE`, `CPI`, `CPP`

---

### C29 - Performing Responsibility

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Description:** Reference of the (single primary) person responsible for performing and completing this task/milestone. The code values should be the Standard ID and the description contain the name.

**Examples:** `GLA`, `CDF`, `MILLER`

---

### C30 - Responsible Organization

**CSPR Requirement:** Mandatory by Process

**Maintained By:** Program Focal

**Description:** Identifies the organization that has responsibility for the task (either supplier, customer, Boeing (optional), production, unauthorized task, schedule margin, etc.)

**Examples:** `NGC`, `BAE`, `CUST`, `EXT`, `UA`, `MARGIN`, `MISC`

**References:** IPSG-9000-002 Allowable Non-Resource Loaded IMS Activity Types and IPSG-9000-008 Integrated Scheduling Requirements Tailoring Matrix

---

### C31 - Risk Interface

**CSPR Requirement:** Mandatory by Process

**Maintained By:** Program Focal

**Description:** Used to identify the Risk # (from Risk Registry ID, BORIS ID, etc.) for the Mitigation Plan Steps included in the IMS.

**Examples:** `P212`, `15`, `6432`, `SUP123`

**References:** IPSG-9000-012 IMS Content Requirements

---

### C32 - Organizational Impact

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Description:** Organizational Impact identifies the impact to other organizations.

**Examples:** `COMMON`, `CRITICAL & COMMON`, `CRITICAL`

**Note:** BDS Programs with contractual IMS CDRL and/or DFAR clause requirements cannot use to identify Hand-offs or SVTs (see IPSG-9000-006 and IPSG-9000-002)

---

### C33 - Schedule Change Authorization

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Description:** Indicates that the cost class scheduler has analyzed the proposed baseline date changes and has determined that the dates will work in the schedule.

**Examples:** `JAB`, `LEP`

---

### C34 - Assembly Control Code (ACC)

**CSPR Requirement:** Mandatory by Process

**Maintained By:** Program Focal

**Description:** Numeric code assigned to each unit time assembly or major assembly position for release of MLP job paper, parts pick manifest and cost accumulation.

**Examples:** `200`, `201`, `300`

**Required When:** Project is Associated with IGOLD

---

### C35 - Delete Me Flag ⚠️ MANDATORY

**CSPR Requirement:** Mandatory

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Baselined:** Yes

**Description:** Flags activities that are to be zeroed out and removed from Cobra and OPP. **Mandatory for CD processing & Change Control.**

**Predetermined Values:** `YES`, `NO`

---

### C36 - Alternate Breakdown Structure

**CSPR Requirement:** Optional

**Maintained By:** Program Focal

**Description:** Used to identify functionality area to perform the task.

**Examples:** `Engineering`, `Business Supplier`, `MFG`

---

### C37 - Effectivity

**CSPR Requirement:** Mandatory by Process

**Maintained By:** Program Focal

**Description:** Unique Identifier for a given aircraft.

**Examples:** `YA001`, `YA002`

**Required When:** Project is Associated with IGOLD

---

### C38 - Recurring/Non Recurring

**CSPR Requirement:** Mandatory by Process

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** Identifies recurring and non-recurring activities for Estimate to Budget Baseline (EtBB) Best Practice.

**Predetermined Values:** `REC`, `NONREC`

**Required When:** Project is Associated with EtBB

---

### C39 - Hours or Dollars

**CSPR Requirement:** Mandatory by Process

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** Identifies activity costs measured either in hours or dollars for Estimate to Budget Baseline (EtBB) Best Practice.

**Predetermined Values:** `HOURS`, `DOLLARS`

**Required When:** Project is Associated with EtBB

---

### C40 - Level 3 Change Control

**CSPR Requirement:** Mandatory by Process

**Maintained By:** Program Focal (Defined by Program in proper format)

**Description:** Identifies activities by Change Document number that will be integrated with Cobra using the Integration Wizard.

**Required When:** Project is Level 3 Change Control (identified in Cobra)

---

### C41 - Material Code

**CSPR Requirement:** Mandatory by Process

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Baselined:** Yes

**Description:** Identifies Critical material represented in the IMS as 'Critical' (required).

**Predetermined Values:** `Critical` (Only 'Critical' should be used. All others are no longer applicable and should not be selected)

**References:** IPSG-4711-003 'IMS Material Integration'

---

### C42 - Schedule Risk Assessment ⚠️ MANDATORY

**CSPR Requirement:** Mandatory

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Baselined:** Yes

**Description:** The field allows users to assign a schedule risk code to an activity that is used by Risk Ranger to calculate the duration input values used in the schedule risk analysis process.

**Examples:** `DES.L.L`, `REQ.M.L`, `NR`, `IA`, `SYS.M.M`

---

### C43 - Reason Code - Actual Start

**CSPR Requirement:** Mandatory by Process

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Baselined:** Yes

**Description:** CAM Visibility users will have to select a Reason code from the C43 code file when they input an Actual Start date which is later than the Baseline Start date.

**Examples:** `BOEING PREDECESSOR`, `TECHNICAL`, `SUSPENDED OPERATIONS`

---

### C44 - ESGP Gate Closure Decision Milestone

**CSPR Requirement:** Mandatory by Process

**Maintained By:** BDS IP&S Home Office Tools Team / BCA Core (Predetermined)

**Baselined:** Yes

**Description:** This field allows users to identify Gate Closure Decision Milestone. A preset of milestones are provided for use.

**Examples:** `Gate 01`, `Gate 02`, `Gate 03`

**References:** BDS IPSG-9000-009

---

### C45 - Contract/Delivery Order/ECP Name

**CSPR Requirement:** Mandatory by Process

**Maintained By:** Program Focal

**Baselined:** Yes

**Description:** This field is used for reporting. It contains a short description (20-30 characters) of the contract/delivery order/ECP. The complete description can be included in the Code Description.

**Examples:** `ESM 158`, `GSP OPT 2`, `LAIRCM 810 EQ`

---

### C46 - SPA Processed Indicator

**CSPR Requirement:** Optional

**Maintained By:** Program Focal (Defined by Program in proper format)

**Baselined:** Yes

**Description:** Identifies closed WPs impacted by a SPA implemented on the program. These WPs end up disconnected with Cobra & further action may be needed.

**Format:** `SPA_MMMYYYY` (e.g., `SPA_Oct2020`)

---

### C47 - Commitment Schedule Type

**CSPR Requirement:** Mandatory by Process

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Baselined:** Yes

**Description:** Code C47 is Specific to the Commitment Scheduling Team, and their reporting requirement.

**Examples:** `Drawings`, `Procurement`, `Sub-Assembly`, `Qual Testing`, etc.

**References:** BDS Programs - refer to IPSG-9000-024 CSCC OPP INTSCH C47 Code Field

---

### C48 - Reason Code - Actual Finish

**CSPR Requirement:** Mandatory by Process

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Baselined:** Yes

**Description:** CAM Visibility users will have to select a Reason code from the C48 code file when they input an Actual Finish date and the actual duration of the task is greater than the baselined duration.

**Examples:** `BOEING PREDECESSOR`, `TECHNICAL`, `SUSPENDED OPERATIONS`

---

### C49-C50 - Reserved

**CSPR Requirement:** Not available for program use

**Maintained By:** IP&S Home Office Tools Team

**Baselined:** Yes

**Description:** Reserved code "positions" for future use to be defined by IP&S Home Office.

---

### C51-C55 - Program Specific Use (Unrotated)

**CSPR Requirement:** Optional

**Maintained By:** Program Focal (Defined by Program)

**Baselined:** Yes

**Description:** Available for program-specific use. "Unrotated" means values remain visible in standard views.

---

### C56-C99 - Program Specific Use (Rotated)

**CSPR Requirement:** Optional

**Maintained By:** Program Focal (Defined by Program)

**Baselined:** No

**Description:** Available for program-specific use. "Rotated" means values are stored differently and may not display in all views.

---

## Project-Level Code Fields

### C1 - No Longer Used

**Description:** No longer Cobra Program. Handled in CSPR Options. Data populated in Project Level User Field (`C_PRGM`).

---

### C2 - CSPR_C2_CME_Program

**CSPR Requirement:** Optional

**Maintained By:** Program Focal (Defined by Program)

**Description:** Month End Cobra Program

---

### C3 - CSPR_C3_IMICS

**CSPR Requirement:** Mandatory by Process

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** ReMICS Flag

---

### C4 - CSPR_C4_LINE_NUM

**CSPR Requirement:** Optional

**Maintained By:** Program Focal (Defined by Program)

**Description:** Line Number

---

### C5 - CSPR_C5_EFFECTIVITY

**CSPR Requirement:** Optional

**Maintained By:** Program Focal (Defined by Program)

**Description:** Effectivity

---

### C6 - CSPR_C6_MAJMIN_MODEL

**CSPR Requirement:** Optional

**Maintained By:** Program Focal (Defined by Program)

**Description:** Major / Minor Model

---

### C9 - C09_CAL_OPT

**CSPR Requirement:** Optional

**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Description:** Accounting Calendar Options

**Predetermined Values:** `DEFAULT`, `TAPESTRY`

---

### C80-C90 - Schedule Visibility Data (SVD)

**Description:** Used in support of the BGS Schedule Visibility Data (SVD)

#### C80 - SVD_C80_Project_Type
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Predetermined Values:** `DEVELOPMENT`, `GOVERNMENT`, `PRODUCTION`, `SUSTAINMENT`

#### C81 - SVD_C81_Prime_Contractor
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Predetermined Values:** `BCA`, `BDS`, `BGS`

#### C82 - SVD_C82_Beta
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Predetermined Values:** `NO`, `YES`

**Description:** Include in NPD or Do not include in NPD identification

#### C83 - SVD_C83_Platform_Model
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Examples:** `777X`, `APACHE`, `KC-46`, `PAR`, `V22`

#### C84 - SVD_C84_Customer_Reporting
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Predetermined Values:**
- `CDRL-SIMPLE COST MANAGEMENT`
- `CDRL-STREAMLINED PERFORMANCE MANAGEMENT`
- `DFAR - DFAR`
- `INTERNAL REPORTING ONLY-SIMPLE COST`
- `INTERNAL REPORTING ONLY-STREAMLINED`

#### C85 - SVD_C85_Sales_Type
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Predetermined Values:**
- `DCS - Direct Commercial Sale`
- `FMS - Foreign Military Sale`
- `USG - US Government`

#### C86 - SVD_C86_Country
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Examples:** `AUSTRALIA`, `ITALY`, `POLAND`, `USA`

#### C87 - SVD_C87_Agile_In_Use
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Predetermined Values:** `NO`, `YES`

**Description:** Has tasks in an Agile Backlog or Does not have Tasks in an Agile Backlog

#### C88 - SVD_C88_In_a_Multi
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Predetermined Values:** `NO`, `YES`

**Description:** The project is in a multi or The project is not in a multi

#### C89 - SVD_C89_Code_Alpha
**Maintained By:** IP&S Home Office Tools Team (Predetermined)

**Predetermined Values:** `MESA`, `OKLAHOMA CITY`, `PHILADELPHIA`, `SAINT LOUIS`

#### C90 - SVD_C90_Code_Delta
**Maintained By:** IP&S Home Office Tools Team (Defined by Program)

---

## Quick Reference - Mandatory Fields

| Code | Field Name | Requirement | Integration |
|------|-----------|-------------|-------------|
| **C4** | Control Account | Mandatory | Cobra/CSPR |
| **C6** | Cost Class | Mandatory | Cobra/CSPR |
| **C11** | Statement of Work Paragraph | Mandatory | CSPR |
| **C35** | Delete Me Flag | Mandatory | Cobra/CSPR |
| **C42** | Schedule Risk Assessment | Mandatory | Risk Ranger |

---

## Quick Reference - Conditional Mandatory Fields

| Code | Field Name | Required When |
|------|-----------|---------------|
| **C1** | WBS | BDS Programs w/ DFAR using Acumen/JSON |
| **C2** | OBS | BDS Programs w/ DFAR using Acumen/JSON |
| **C20** | Make/Buy | Project Associated with MES |
| **C34** | Assembly Control Code | Project Associated with IGOLD |
| **C37** | Effectivity | Project Associated with IGOLD |
| **C38** | Recurring/Non Recurring | Project Associated with EtBB |
| **C39** | Hours or Dollars | Project Associated with EtBB |
| **C40** | Level 3 Change Control | Project is Level 3 Change Control |

---

## 🏢 Best Practices for Code Field Usage

### 1. Understand Your Program Requirements
- Check which system integrations your program uses (Cobra, CSPR, MES, IGOLD, EtBB)
- Review program-specific IPSG documents referenced in code field descriptions
- Confirm with Program Focal which optional fields are required by your program

### 2. Maintain Data Integrity
- **⚠️ NEVER manually edit C4 (Control Account) code file** - Use OPP/Cobra mapping only
- Follow predetermined value lists for fields maintained by IP&S Home Office Tools Team
- Use consistent naming conventions (e.g., `Last_First_BEMS` format for personnel fields)

### 3. Baseline Management
- Fields marked "Baselined: Yes" should be set before creating baseline
- Changes to baselined fields after baselining may require change control
- Document rationale for changes to baselined code field values

### 4. Code File Structure
- **Punctuated Significant:** Hierarchical with delimiters (e.g., `1.3.1`)
- **Fixed Format:** Non-hierarchical flat list (e.g., `HAA10AAL`)
- Choose format based on program's organizational structure

### 5. VBA Access to Code Fields

**Accessing Code Values:**
```vb
' Get code value for an activity
Dim act As Object
Set act = proj.Activities.Item("ACT001")
Dim wbsCode As String
wbsCode = act.GetField("C1")  ' Returns WBS code

' Get code description
Dim wbsDesc As String
wbsDesc = act.GetField("C1.DESCRIPTION")

' Set code value
act.SetField "C1", "1.3.1"
```

**Iterating Through Code Files:**
```vb
' Loop through all values in a code file
Dim codeDir As Object
Set codeDir = proj.CodeDirectories.Item("C1")

Dim code As Object
For Each code In codeDir.Codes
    Debug.Print code.Code & " - " & code.Description
Next code
```

### 6. Import/Export Considerations

**Transfer.dat Script Example:**
```
# Export activities with critical code fields
EXPORT csv ActivityExport
TABLE ACT
FIELD ACT_ID DESCRIPTN C1 C2 C4 C6
FIELD ESDATE EFDATE ASDATE AFDATE
```

**⚠️ CRITICAL:** Remember Transfer.dat commands must be UPPERCASE

### 7. Calculated Field Access

**Using Code Fields in Calculated Fields:**
```
# Display Control Account description
C4.DESCRIPTION

# Check if WBS is at specific level
IIF(LEVEL(C1) >= 3, "Detailed", "Summary")

# Concatenate multiple codes
C1 + " / " + C2
```

---

## 🐛 Common Issues and Solutions

### Issue: Code field not populating from Cobra

**Cause:** Manual edit of C4 (Control Account) code file or mapping not run

**Fix:**
1. Do NOT manually edit C4 code file
2. Use OPP/Cobra mapping tool to refresh code files
3. Verify Cobra control accounts are flagged for OPP
4. Contact IP&S Home Office Tools Team if mapping fails

### Issue: "Code value not found" error

**Cause:** Activity assigned code value that doesn't exist in code file

**Fix:**
1. Check code file (Tools > Codes) to verify value exists
2. Ensure code value spelling/format matches exactly
3. For predetermined fields, use only approved values
4. Refresh code files from Cobra if applicable

### Issue: Baseline includes wrong code values

**Cause:** Code fields updated after baseline created

**Fix:**
1. Review Change Control process requirements
2. Update code values BEFORE creating new baseline
3. Document changes per program baseline change process
4. For fields marked "Baselined: Yes", changes may require formal approval

### Issue: Integration failure with external system

**Cause:** Missing mandatory code field or incorrect format

**Fix:**
1. Review Quick Reference - Mandatory Fields table above
2. Verify all conditional mandatory fields for your program's integrations
3. Check predetermined value lists match exactly
4. Validate format (punctuated vs fixed) matches system requirements

---

## 📚 Related Boeing Documents

- **IPSG-9000-002:** Allowable Non-Resource Loaded IMS Activity Types
- **IPSG-9000-003:** Integrated Scheduler Code Field
- **IPSG-9000-006:** [Hand-offs and SVTs]
- **IPSG-9000-007:** WBS, OBS, and CAM Fields
- **IPSG-9000-008:** Integrated Scheduling Requirements Tailoring Matrix
- **IPSG-9000-009:** [ESGP Gate Closure Decision Milestones]
- **IPSG-9000-012:** IMS Content Requirements
- **IPSG-9000-019:** Quantifiable Backup Data
- **IPSG-9000-024:** CSCC OPP INTSCH C47 Code Field
- **IPSG-4711-003:** IMS Material Integration

---

*For VBA automation, see [VBA-API-Reference.md](VBA-API-Reference.md)*
*For critical warnings, see [Critical-Warnings-and-Patterns.md](Critical-Warnings-and-Patterns.md)*

**Last Updated:** 2025-10-28
**Source:** CSPR_OPP_DataDictionary.csv (Boeing IP&S Home Office Tools Team)
