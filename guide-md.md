---
name: guide
description: Generates a comprehensive, user-centric `USER_GUIDE.md` that explains *how to use* the application, interpret its interface, and resolve common errors.
---

## Phase 1: Code Analysis (The "User Journey" Scan)

Before writing a single word, the LLM must analyze the entry points of the code (`main.py`, `App.js`, `index.html`) to construct the User Journey, The User will only get the .exe file.

**The LLM must identify:**

1.  **Inputs & Actions:**
    * What does the user click? (Button labels, Menu items)
    * What does the user type? (Input fields, Arguments)
    * What files does the user upload? (Formats: `.csv`, `.json`, images)

2.  **Logic & Constraints (The "Rules"):**
    * Look for `if` statements validating inputs (e.g., `if age < 18`).
    * Look for regex patterns (e.g., Email validation, Password strength).
    * Look for file constraints (e.g., "Max 5MB", "Must include header row").

3.  **Outputs:**
    * What happens when successful? (Success message, File download, Chart rendered).

4.  **Error Handling:**
    * Scan `try/except` blocks (Python) or `.catch()` blocks (JS).
    * Map the *error message string* in the code to the *cause* of the error.

---

## Phase 2: Drafting the User Guide

The generated documentation **MUST** follow this structure. Do not output generic text; anchor every instruction to actual code features.

### Section 1: Application Overview
* **What is it?** (1 sentence summary)
* **Who is it for?** (Target audience)
* **Key Features:** (Bullet points of capabilities)

### Section 2: Getting Started (The "Happy Path")
* **Prerequisites:** (API Keys needed? `.env` setup? Internet connection?)
* **Launch:** How to open the app (e.g., "Double click the `.exe`" or "Open localhost:3000").

### Section 3: Step-by-Step Usage Guide
* Break down the workflow into numbered steps.
* **Format:** `Action` -> `Result`.
* *Example:* "Step 1: Click 'Upload'. Select a `.csv` file. The table will populate with data."

### Section 4: Interface & Controls (Deep Dive)
* List every visible field/button.
* Explain its purpose.
* *Example:*
    * **Field:** "Threshold Slider"
    * **Purpose:** Filters out results below this value.
    * **Tip:** Keep between 0.5 - 0.8 for best results.

---

## Phase 3: Validation & Troubleshooting (Crucial)

This section translates code logic into user help. The LLM must explicitly map constraints found in the code to user solutions.

**Format for the Output:**

| Error / Warning Message | Possible Cause | How to Fix |
| :--- | :--- | :--- |
| *(Extract exact string from code)* | *(Explain the logic check)* | *(User action)* |

**Example Mapping:**
* *Code:* `if file.size > 5000000: raise Error("File too big")`
* *Doc Entry:* **Error:** "File too big" | **Cause:** Upload exceeded 5MB | **Fix:** Compress your file or split it.

---

## Phase 4: Output Template

The LLM must generate the final Markdown using this exact template:

```markdown
# [App Name] User Guide

## Introduction
[Brief description of what the app does]

## How to Use

### 1. Launching the App
[Instructions on running the EXE or Localhost]

### 2. The Workflow (Step-by-Step)
1. **[Action Name]**: [Instruction]
2. **[Action Name]**: [Instruction]
   - *Note: [Relevant constraint found in code]*

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| [Field Name] | [What it does] | [e.g., YYYY-MM-DD] |

## Troubleshooting & Validations

If you see an error, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| [Error String] | [Logic Explanation] | [Correction Step] |

```

---

## Phase 5: Final Review

Before finalizing the README, the LLM must verify:

1. Are all **API Keys** or **Environment Variables** mentioned?
2. Is the **File Format** for uploads explicitly defined (e.g., "CSV must have headers: Name, Email")?
3. Are **Output Locations** specified? (e.g., "Reports are saved in the /downloads folder").

```

```