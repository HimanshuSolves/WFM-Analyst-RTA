📁 Workbook Structure

---

## 📘 WFM - ANALYST FUNCTIONS.xlsm

### 1. WFM ANALYSIS
- Primary KPI reference and summary table.
- Displays metrics such as AHT (Average Handle Time), Occupancy, etc.
- Includes formula columns (likely used for learning/reference rather than calculations).

### 2. WFM FORMULAS
- A formula reference table with Excel-calculable expressions.
- Explains WFM-specific metrics like Shrinkage, Utilization, etc.
- Designed for training, auditing, or development support.

### 3. WFM CONCEPTS & TERMINOLOGY
- Glossary-style sheet defining key WFM terms.
- Examples include: Adherence, FTE, Absenteeism, etc.
- Ideal for onboarding or documentation.

### Macros and Automation
- ⚠ No macros detected in this workbook.
- Workbook likely used as a reference/support tool rather than an automated system.

---

## 📗 WFM - REAL TIME ANALYST.xlsm

### 1. ACCESS SHEETS
- Control panel for toggling visibility of other sheets.
- Contains form control checkboxes (e.g., C10, F10, I10, L10).
- Connected to macros that show/hide sheets like:
  - WFM RAW DATA
  - BACKGROUND CALCULATIONS
  - LOGIN HOURS CALCULATOR

### 2. WFM PERFORMANCE ANALYSIS
- Main dashboard or pivot-driven sheet.
- Contains summary metrics such as Productivity, SLA, AHT, etc.
- Driven by pivot tables based on raw and calculated data.

### 3. WFM RAW DATA
- Source sheet for performance calculations.
- Includes columns like Agent ID, Date, AHT, Break Duration, etc.
- Estimated: 300–500 rows × 50+ columns.

### 4. BACKGROUND CALCULATIONS
- Processes raw data into structured formats.
- Used for formulas related to login/logout, adherence, productivity.

### 5. LOGIN HOURS CALCULATOR
- Utility to calculate total login duration.
- Likely includes: Agent Name, Login/Logout, Breaks, Net Hours.

---

### Macros and Automation

- Stored in: `vbaProject.bin`
- Checkbox-driven sheet visibility macros.
- Include:
  - `UpdateDashboard()`
  - `ToggleVisibility()`
- Example:
```vba
If Range("C10").Value = True Then
    Sheets("WFM RAW DATA").Visible = xlSheetVisible
Else
    Sheets("WFM RAW DATA").Visible = xlSheetVeryHidden
End If
