# Copy Excel File Sheet to Another Excel File

A simple Python-based tool that copies all sheets from a **Source Excel File** to a **Destination Excel File**.

---

## Instructions

### Step 1: Install Required Libraries
Run the provided installer script to automatically install the required packages:
```bash
python install_requirements.py
```
### Step 2: Run the Script
Run the provided installer script to automatically install the required packages:
```bash
python copy_sheets_to_excel.py
```

---
---

## How to Use (For End Users)

1. Navigate to the [`copy_sheets_to_excel.exe`](./dist/copy_sheets_to_excel.exe) file.
2. **Double-click to launch.**
   - First, select the **Destination Excel file** (the file to copy sheets **into**).
   - Next, select the **Source Excel file** (the file to copy sheets **from**).
3. The tool will:
   - Copy all sheets from the source to the destination
   - Rename any duplicate sheet names (e.g., `Sheet1_copy1`, `Sheet1_copy2`, etc.)
   - Save the updated destination file
4. A terminal window will display confirmation messages.
