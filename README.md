# SFA Fencing Camp Schedule Generator

This tool automates the scheduling of lessons for the fencing camp. It uses a two-step process to allow for manual adjustments in Excel before generating final reports.

## Workflow

### Step 1: Initial Generation
Run this command to read the student data from `Aug_shenzhen.xlsx` and generate the initial master schedule.

```bash
python3 schedule_generator.py --step1
```

**Outputs:**
- `SummerCamp_Schedule.xlsx`: The master dashboard with one tab per day (Day 1-6).
- `schedule.json`: Data for the web dashboard.
- `schedule_summary.csv`: A quick summary of assignments and shortages.

---

### Step 2: Manual Adjustments (Optional)
Open `SummerCamp_Schedule.xlsx` and make any manual changes you like (e.g., swapping students between coaches or moving them to different time slots). 

**Rules for editing:**
- Only change the student names in the coach columns.
- Do not change the coach names in the headers.
- Do not change the time slot labels in the first column.

---

### Step 3: Final Reports & Validation
Run this command to read the (potentially edited) `SummerCamp_Schedule.xlsx` and generate the final coach-centric and kid-centric reports.

```bash
python3 schedule_generator.py --step2
```

**Outputs:**
- `schedule_by_coach.xlsx`: A professional report with **one tab per coach**. columns=Days, rows=Time.
- `schedule_by_kid.csv`: A simple list of sessions per student.
- `schedule.json`: Updates the web dashboard to reflect your manual edits.

**Validation:**
Step 2 automatically runs a check to ensure that the total number of assigned lessons matches the original requests in `Aug_shenzhen.xlsx`. It will alert you to any "SHORT" or "OVER" assignments.

---

## Web Dashboard
Open `index.html` in your browser to view the interactive schedule.
- **Search**: Find any student's schedule quickly.
- **Edit Mode**: (Password: `admin123`) Move students directly in the web UI. If you do this, use the "Export JSON" button to save your changes back to the project.
