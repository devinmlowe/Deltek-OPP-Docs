# Deltek Open Plan Documentation - README

This directory contains optimized documentation for **Deltek Open Plan® Professional** OLE Automation, organized as a modular knowledge base for Custom GPT applications.

---

## 📁 File Structure

### Reference Files (Attach to Custom GPT Knowledge Base)

| File | Purpose | Token Count | Topics Covered |
|------|---------|-------------|----------------|
| **[Critical-Warnings-and-Patterns.md](Critical-Warnings-and-Patterns.md)** | Fatal mistakes & required patterns | ~2,000 | VBA fatal mistakes, Transfer.dat case sensitivity, field naming, pre-flight checklists, error messages, anti-patterns |
| **[VBA-API-Reference.md](VBA-API-Reference.md)** | Complete VBA OLE Automation API | ~13,000 | 33+ objects, properties, methods, collections, database field names, common patterns, inline debugging |
| **[Import-Export-Reference.md](Import-Export-Reference.md)** | Transfer.dat scripting language | ~11,000 | Commands, table types, XML formatting, default scripts, inline debugging |
| **[Calculated-Fields-Reference.md](Calculated-Fields-Reference.md)** | Expression language for formulas | ~9,000 | 50+ functions, operators, constants, user variables, formula patterns, inline debugging |

**Total Token Count:** ~35,000 tokens (86% reduction from original 250k token developer guide)

### Configuration Files

| File | Purpose |
|------|---------|
| **[SYSTEM-PROMPT-Template.md](SYSTEM-PROMPT-Template.md)** | Template for Custom GPT system instructions |
| **README-Documentation.md** | This file - explains structure and usage |

### Source & Archive Files

| File | Purpose |
|------|---------|
| **[DeltekOpenPlanDeveloperGuide.md](DeltekOpenPlanDeveloperGuide.md)** | Original 27,161-line developer guide |
| **[DeltekOpenPlanDeveloperGuide-LLM-Optimized.md](DeltekOpenPlanDeveloperGuide-LLM-Optimized.md)** | Ultra-compact version (8k tokens) |
| **_ARCHIVE_DeltekOpenPlanDeveloperGuide-LLM-Reference-v1-SingleFile.md** | Previous single-file version (archived) |

---

## 🚀 Quick Start - Setting Up Your Custom GPT

### Step 1: Create Custom GPT in OpenAI

1. Go to [ChatGPT](https://chat.openai.com) → Your Profile → My GPTs
2. Click **"Create a GPT"**
3. Name it: **"Deltek Open Plan Assistant"** (or your preference)

### Step 2: Configure Instructions (System Prompt)

1. Open [SYSTEM-PROMPT-Template.md](SYSTEM-PROMPT-Template.md)
2. Copy the entire prompt content
3. Paste into the **"Instructions"** field in your Custom GPT
4. Customize if needed (tone, additional rules, etc.)

### Step 3: Upload Knowledge Files

In the **"Knowledge"** section, upload these four files:

1. ✅ `Critical-Warnings-and-Patterns.md`
2. ✅ `VBA-API-Reference.md`
3. ✅ `Import-Export-Reference.md`
4. ✅ `Calculated-Fields-Reference.md`

**Important:** All four files will be loaded into context at conversation start (~35k tokens total).

### Step 4: Test Your Assistant

Try these test prompts:

**Test 1 - VBA:**
> "Create a VBA script that loops through all activities and sets USER1 to 'Reviewed' for completed tasks."

**Test 2 - Import/Export:**
> "Create a Transfer.dat script to export activities with their resources to CSV."

**Test 3 - Calculated Fields:**
> "Create a calculated field that shows 'Late' if the activity's actual finish is after its early finish."

**Test 4 - Debugging:**
> "My VBA script runs but the dates aren't updating. What's wrong?"

---

## 📊 Design Rationale

### Why Modular Files?

**Original Challenge:** 27,161-line developer guide = 250k tokens (too large for most contexts)

**Solution Evolution:**
1. **First Optimization:** Single ultra-compact file (8k tokens) - too sparse for code generation
2. **Second Enhancement:** Single reference file with examples (28k tokens) - better but not optimal for RAG
3. **Current Structure:** Four focused files (35k tokens total) - **best balance**

**Benefits of Modular Approach:**

| Benefit | Description |
|---------|-------------|
| **Better Semantic Boundaries** | Each file covers one topic, easier for RAG retrieval |
| **Inline Context** | Debugging tips stay with relevant content |
| **Independent Maintenance** | Update VBA docs without touching Import/Export |
| **Topic Isolation** | Users typically ask about one topic per conversation |
| **Future-Proof** | Can optimize individual files without restructuring |

### Why ~35k Tokens (Not Less)?

- **Critical Warnings** (~2k): Prevents 90% of common mistakes
- **Complete API** (~13k): Covers all 33+ objects with examples
- **All Commands** (~11k): Complete Transfer.dat reference with XML
- **All Functions** (~9k): Complete calculated field function library

**Trade-off:** Slightly higher token usage, but:
- ✅ Provides complete, accurate code examples
- ✅ Reduces back-and-forth clarifications
- ✅ Includes inline debugging for common issues
- ✅ Still saves 215k tokens (86%) vs original

---

## 🎯 How the System Works

### At Conversation Start

When a user starts a new conversation with your Custom GPT:

1. **System Prompt Loads:** The instructions from `SYSTEM-PROMPT-Template.md` are active
2. **All Knowledge Files Load:** All four reference files (~35k tokens) load into context
3. **User Query:** User asks a question
4. **LLM Processing:**
   - Checks critical rules from system prompt (memorized)
   - Retrieves relevant details from appropriate reference file(s)
   - Generates response with inline warnings and complete code

### Conversation Flow Example

**User:** "Create a VBA script to export activities"

**LLM Internal Process:**
1. System prompt: "Check VBA critical rules"
2. Reference: `VBA-API-Reference.md` for object syntax
3. Reference: `Import-Export-Reference.md` for alternative approaches
4. Generate: Complete VBA script with `.Login()`, `.TimeAnalyze()`, `.Save()`
5. Include: Pre-flight checklist and debugging tips

**Response Contains:**
- Critical warnings mentioned upfront
- Complete, ready-to-run code
- Inline comments explaining each step
- Debugging tips for common issues

---

## 🔍 Content Organization

### Critical-Warnings-and-Patterns.md

**Structure:**
```
├── ⚠️ Fatal Mistakes
│   ├── VBA Automation (3 fatal mistakes)
│   ├── Import/Export Scripts (2 fatal mistakes)
│   └── Calculated Fields (2 fatal mistakes)
├── ✅ Required Patterns
│   ├── VBA minimum viable script
│   ├── Transfer.dat minimum script
│   └── Calculated field syntax
├── ❌ Anti-Patterns (with corrections)
├── Pre-Flight Checklists (before running code)
├── Common Error Messages (with solutions)
└── Performance Tips
```

### VBA-API-Reference.md

**Structure:**
```
├── Quick Reference (common operations)
├── Core Objects (Application, Project, Activity)
├── Collection Objects (Activities, Resources, etc.)
├── View Objects (Barchart, Network, Spreadsheet)
├── File Cabinet Objects
├── Other Objects (25+ additional objects)
├── Common Patterns (filtering, sorting, adding)
├── Enumerations & Constants
└── Database Field Names (complete list)
```

**Each Object Includes:**
- Description and purpose
- Key properties with data types
- Key methods with parameters
- Code examples
- 🐛 Debugging Tips section
- ✅ Best Practices section

### Import-Export-Reference.md

**Structure:**
```
├── Transfer.dat Overview
├── ⚠️ CRITICAL: UPPERCASE Warning
├── Default Scripts (15+ pre-built scripts)
├── Core Commands (IMPORT, EXPORT, TABLE, FIELD, etc.)
├── Additional Commands (20+ commands)
├── Table Types Reference
├── XML Commands
├── Complete XML Examples
├── 🐛 Debugging Guide
└── ✅ Best Practices Summary
```

### Calculated-Fields-Reference.md

**Structure:**
```
├── Overview & Elements
├── Constants (dates, durations, text, enums)
├── ⚠️ Field Name Warning
├── Operators (mathematical, relational, logical)
├── User-Defined Variables
├── Functions (50+ functions)
│   ├── Mathematical
│   ├── Date/Time
│   ├── String Manipulation
│   ├── Logical
│   ├── Data Access
│   └── Hierarchy Functions
├── 🐛 Debugging Guide
├── ✅ Best Practices
└── Common Formula Patterns
```

---

## 🛠️ Maintenance Guide

### When to Update Files

**Update Critical-Warnings-and-Patterns.md when:**
- New fatal mistakes are discovered
- Common error patterns emerge from user feedback
- Anti-patterns need additional examples

**Update VBA-API-Reference.md when:**
- New objects/properties/methods are documented
- Better code examples are developed
- Additional debugging scenarios are identified

**Update Import-Export-Reference.md when:**
- New Transfer.dat commands are discovered
- Additional default scripts are created
- XML formatting examples need expansion

**Update Calculated-Fields-Reference.md when:**
- New functions are documented
- Additional formula patterns are developed
- Function syntax needs clarification

### How to Update Files

1. **Edit the specific file** - Don't need to regenerate all files
2. **Maintain token budget** - Keep files under their target sizes
3. **Preserve structure** - Keep visual markers (⚠️, ✅, ❌, 🐛)
4. **Update cross-references** - Ensure links between files are accurate
5. **Test with Custom GPT** - Verify the changes improve responses
6. **Update this README** - Reflect any structural changes

### Version Control

This documentation uses Git for version control:

```bash
# Make your edits, then commit
git add docs/
git commit -m "Update VBA reference with new examples"
git push origin main
```

### Token Count Monitoring

Check token counts if files grow significantly:

```bash
# Use an LLM token counter or estimate: ~4 characters = 1 token
wc -c Critical-Warnings-and-Patterns.md
# Divide by 4 for rough token estimate
```

**Target Token Counts:**
- Critical-Warnings: ~2,000 tokens (max 3,000)
- VBA-API: ~13,000 tokens (max 15,000)
- Import-Export: ~11,000 tokens (max 13,000)
- Calculated-Fields: ~9,000 tokens (max 11,000)

**Total Budget:** Keep under 40,000 tokens total

---

## 🔗 Related Documentation

### In This Repository

- **[Developer Guide Overview](developer-guide/README.md)** - Networked documentation structure
- **[Getting Started](developer-guide/getting-started/overview.md)** - Introduction for new users
- **[Object Hierarchy](developer-guide/object-hierarchy/README.md)** - Visual object relationships
- **[Objects Reference](developer-guide/objects/README.md)** - Individual object documentation
- **[Code Examples](developer-guide/examples/README.md)** - Complete working examples

### External Resources

- **Deltek Open Plan Product Page:** [Deltek Website](https://www.deltek.com)
- **Official Support:** Contact Deltek for technical support
- **Community Forums:** Check Deltek community for user discussions

---

## 📝 Changelog

### Version 2.0 - Modular RAG-Optimized Structure (Current)

**Date:** 2025-10-28

**Changes:**
- Split single LLM-Reference file into four focused files
- Added inline debugging tips with each section
- Created system prompt template for Custom GPT
- Organized by topic for better semantic retrieval
- Added visual markers (⚠️, ✅, ❌, 🐛) for pattern recognition
- Included pre-flight checklists and anti-patterns
- Total: ~35k tokens across 4 files

### Version 1.0 - Single LLM-Reference File (Archived)

**Date:** [Previous Date]

**Changes:**
- Consolidated developer guide into single reference file
- Added VBA API documentation
- Added Import/Export Transfer.dat documentation
- Added Calculated Fields documentation
- Total: ~28k tokens in 1 file

**Archive Location:** `_ARCHIVE_DeltekOpenPlanDeveloperGuide-LLM-Reference-v1-SingleFile.md`

---

## 💡 Tips for Best Results

### For Custom GPT Creators

1. **Copy the system prompt exactly** - The critical rules section is carefully crafted
2. **Upload all four files** - They work together as a complete reference
3. **Test with real scenarios** - Use actual user questions from your workflow
4. **Refine the tone** - Adjust system prompt personality to match your needs
5. **Monitor token usage** - Watch for context window issues with long conversations

### For End Users (Using the Custom GPT)

1. **Be specific** - "Create VBA to update USER1" is better than "help with VBA"
2. **Mention your version** - "Open Plan 8.x" vs "Open Plan 7.x" can affect code
3. **Include error messages** - Exact error text helps with debugging
4. **Ask for explanations** - "Why do I need TimeAnalyze?" helps you learn
5. **Request modifications** - "Add error handling" or "Make it work for filtered activities"

### For Developers Maintaining This Repo

1. **Keep files focused** - Don't let scope creep between files
2. **Preserve examples** - Working code examples are the most valuable content
3. **Update inline** - Add debugging tips right where relevant, not in separate sections
4. **Test additions** - Verify new examples actually work in Open Plan
5. **Document edge cases** - Unusual scenarios are where users need most help

---

## ❓ FAQ

### Q: Can I use these files outside of Custom GPT?

**A:** Yes! These are standard markdown files. You can:
- Use them in any LLM with file upload (Claude, Gemini, etc.)
- Read them directly as reference documentation
- Convert them to HTML for a knowledge base website
- Import them into your own documentation system

### Q: Why not just use the original developer guide?

**A:** The original guide is comprehensive but:
- 27,161 lines = 250k tokens (too large for most AI contexts)
- Not optimized for LLM retrieval (narrative structure vs reference structure)
- Lacks inline debugging tips for common issues
- Doesn't emphasize fatal mistakes prominently

### Q: Can I merge the files back into a single file?

**A:** Yes! Simply concatenate them:

```bash
cat Critical-Warnings-and-Patterns.md \
    VBA-API-Reference.md \
    Import-Export-Reference.md \
    Calculated-Fields-Reference.md \
    > Combined-Reference.md
```

Or use the archived single file version.

### Q: How do I add a new object to VBA-API-Reference.md?

**A:** Follow the existing pattern:

```markdown
### OPYourNewObject

**Description:** Brief description of what this object represents.

**Access:** How to get this object (e.g., `project.YourNewObjects`)

#### Key Properties

| Property | Type | Description |
|----------|------|-------------|
| PropertyName | String | What it represents |

#### Key Methods

| Method | Returns | Description |
|--------|---------|-------------|
| MethodName(params) | Object | What it does |

#### Example

\`\`\`vb
' Example code here
\`\`\`

#### 🐛 Debugging Tips

- Tip 1
- Tip 2

#### ✅ Best Practices

- Practice 1
- Practice 2
```

### Q: What if I find an error in the documentation?

**A:** Please:
1. Create an issue in this GitHub repository
2. Include the file name and section
3. Describe the error and suggested correction
4. Include a code example if applicable

### Q: Can I customize the system prompt for my specific use case?

**A:** Absolutely! The template is a starting point. Common customizations:
- Add your company's coding standards
- Include project-specific field names or conventions
- Adjust tone (more formal, more casual)
- Add restrictions (e.g., "Only provide VBA, not Python alternatives")

---

## 📄 License

This documentation is based on the Deltek Open Plan® Professional Developer Guide.

**Deltek Open Plan® Professional** is a registered trademark of Deltek, Inc.

This documentation is provided for reference purposes. Always refer to official Deltek documentation for authoritative information.

---

## 📧 Contact

For questions about this documentation structure or repository:
- **GitHub Issues:** [Create an issue](https://github.com/your-repo/issues)
- **Email:** [Your contact email if applicable]

For Deltek Open Plan product support:
- **Deltek Support:** Contact through your Deltek support portal

---

**Last Updated:** 2025-10-28
**Documentation Version:** 2.0 (Modular RAG-Optimized)
**Repository:** https://github.com/devinmlowe/Deltek-OPP-Docs
