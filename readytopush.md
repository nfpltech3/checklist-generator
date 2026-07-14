---
name: readytopush
description: A structured Antigravity IDE workflow that ensures Python and JavaScript projects are clean, secure, dependency-accurate, build-tested, iteratively validated, and fully ready for GitHub deployment.
---

# 1. Pre-Push Validation Workflow (Mandatory Before GitHub Push)

Before pushing, follow this structured process inside Antigravity IDE:

### Step 1 – 🛡️ MANDATORY ISOLATION CHECK (Python)

* **PROTOCOL:** The AI must never execute a "raw" python command.
* **RULE:** Every Python command (pip, python, pyinstaller) **MUST** be chained with the venv activation command.
* **Fail Safe:** If `venv` does not exist → Create it immediately.

### Step 2 – JavaScript Environment

* Remove old `node_modules`
* Run fresh `npm install`
* Ensure lockfile consistency

---

### Step 3 – Compile / Build / Run to Detect Errors

Before pushing:

* Run the application (Ensure `venv` is active!)
* Build it (if applicable) - **Strictly use .spec file** for Python builds to ensure assets (logos/icons) are bundled correctly.
* Fix syntax/runtime/import errors
* Repeat until clean execution

This must be iterative:
**Activate Venv → Run → Fix → Re-run → Fix → Validate**

Do not push untested code.

---

# 2. Required Repository Files

| File | Purpose |
| --- | --- |
| `README.md` | Setup and usage instructions |
| `.gitignore` | Prevents secrets and system files |
| `requirements.txt` | **VENV-ONLY** Python dependencies |
| `package.json` | JS dependencies & scripts |
| `.env.example` | Environment template |
| `*.spec` | PyInstaller Build Configuration |

---

# 3. Python Workflow (⚠️ STRICT VENV POLICY ⚠️)

## Step 1 – Create Virtual Environment (If Missing)

You cannot skip this.

```bash
python -m venv venv

```

## Step 2 – ACTIVATE Virtual Environment

**You must perform this step every time you open a new terminal.**

**Windows:**

```cmd
venv\Scripts\activate

```

**Mac/Linux:**

```bash
source venv/bin/activate

```

**✅ VERIFICATION:**
Your terminal line must look like this: `(venv) C:\path\to\project>`

---

## Step 3 – Install Only Required Packages

**WARNING:** Do not run this command unless `(venv)` is visible.

1. Inspect project imports.
2. Install only required packages.

Example:

```bash
pip install flask requests pandas

```

---

## Step 4 – Validate Code by Running It

Run main entry file (Ensure `venv` is active):

```bash
python main.py

```

Fix:

* Missing imports
* Module errors
* Syntax issues
* Runtime exceptions

Repeat until clean.

---

## Step 5 – Freeze Accurate Dependencies

Only after successful execution **inside the venv**:

```bash
pip freeze > requirements.txt

```

**Why?** This ensures global system packages are NOT included in your project requirements.

---

## Step 6 – Build / Compile Validation

### For Backend / CLI Apps

Run directly:

```bash
python main.py

```

Ensure no runtime errors.

### For Python Desktop Applications (Tkinter / PyQt / etc.)

**CRITICAL PREREQUISITE:**
Ensure your code handles absolute paths for assets using `sys._MEIPASS`. Standard relative paths will fail inside the EXE.

#### 1. Generate Spec File (First Time Only)

Do not use this command to build. Use it only to generate the config file:

```bash
pyinstaller --name="MyApp" --onefile --windowed --icon="assets/icon.ico" main.py

```

#### 2. Configure Assets in `MyApp.spec`

Open the generated `.spec` file. Find the `datas=[]` list and map your folders/files:

```python
a = Analysis(
    ...
    datas=[
        ('assets/logo.png', 'assets'),  # ('Source Path', 'Dest Folder in Exe')
        ('config.json', '.')
    ],
    ...
)

```

#### 3. Build the Executable

Always build using the spec file to ensure assets are included:

```bash
pyinstaller MyApp.spec

```

#### 4. Verify

* Check `dist/` folder for the executable.
* Run it to ensure **logos and icons load correctly**.
* Ensure no console window appears (if GUI).

Only push if the Spec file build works perfectly.

---

# 4. JavaScript Workflow (Node.js)

## Step 1 – Fresh Install

```bash
rm -rf node_modules
npm install

```

## Step 2 – Run Project

```bash
npm start

```

or

```bash
npm run dev

```

Fix:

* Missing packages
* Incorrect scripts
* Runtime errors

## Step 3 – Production Build (If Applicable)

For frontend apps:

```bash
npm run build

```

Ensure:

* No build warnings breaking output
* No unresolved imports
* No missing environment variables

Fix iteratively.

---

# 5. README.md Structure (Updated)

README must reflect:

* **EXPLICIT INSTRUCTIONS TO ACTIVATE VENV**
* Only project-specific dependency installation
* Build and compile steps
* **Spec file usage** for EXE generation

## Recommended README Template

```markdown
# Project Name

Short description.

## Tech Stack
- Python 3.11 / Node 20
- Framework used

---

## Installation

### Clone
git clone https://github.com/username/project-name.git
cd project-name

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create virtual environment
python -m venv venv

2. Activate (REQUIRED)

Windows:
venv\Scripts\activate

Mac/Linux:
source venv/bin/activate

3. Install dependencies
pip install -r requirements.txt

4. Run application
python main.py

---

### Build Executable (For Desktop Apps)

1. Install PyInstaller (Inside venv):
pip install pyinstaller

2. Build using the included Spec file (Ensure you do not run main.py directly):
pyinstaller MyApp.spec

3. Locate Executable:
The application will be generated in the `dist/` folder.

---

## JavaScript Setup

1. Install dependencies
npm install

2. Run project
npm start

3. Production build (if applicable)
npm run build

---

## Environment Variables

Copy:
cp .env.example .env

Add required values.

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit venv or node_modules.
- Run and test before pushing.

```

---

# 6. .gitignore

## Python

```gitignore
venv/
__pycache__/
*.pyc
.env
dist/
build/
# *.spec  <-- REMOVED: Keep .spec files in repo for consistent builds!

.vscode/
.idea/

.DS_Store
Thumbs.db

```

## JavaScript

```gitignore
node_modules/
.env
dist/
build/

.vscode/
.idea/

.DS_Store
Thumbs.db

```

---

# 7. Final Git Push Workflow

After:

* Clean execution **(CONFIRM VENV WAS USED)**
* Successful build (via Spec file)
* Accurate dependencies **(GENERATED FROM VENV)**
* Verified README

Run:

```bash
git init
git add .
git commit -m "Initial validated commit"
git remote add origin https://github.com/username/repo-name.git
git branch -M main
git push -u origin main

```

---

# 8. Language-Specific Considerations

## Python (ZERO TOLERANCE POLICY)

* **ALWAYS USE VENV.**
* **NEVER INSTALL GLOBALLY.**
* If `(venv)` is not in your terminal, STOP.
* Install only required imports.
* Freeze dependencies after validation.
* **Always use `.spec` files** for PyInstaller builds.
* Rebuild if errors occur.

## JavaScript

* Always test production build.
* Never commit node_modules.
* Validate scripts in package.json.
* Ensure lockfile is included.