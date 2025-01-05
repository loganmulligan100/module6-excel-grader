# Module 6 Excel Grader

This is a Python script designed to automatically grade most parts of the Module 6 Excel assignment. It checks for correct labels, bold formatting, bottom borders, proper formulas (with cell ranges instead of direct additions), and at least some Accounting Number formatting. It also handles small row offsets (like if a student inserted an extra row for Income or Expenses).

> **Note:** It does **not** grade the chart portion of the assignment because the library used for parsing Excel files (openpyxl) does not reliably read existing charts from a saved `.xlsx` file.

---

## Features

- **Label & Format Checks**  
  Verifies that required cells (like "Income," "Total Income," "Expenses," etc) exist in the correct location or are offset by one row. Checks if certain cells are bold or non-bold as specified.

- **Formula Checks**  
  Ensures students used cell-range-based formulas (like `=SUM(B5:B6)`) rather than direct additions (like `=B5+B6`). If it detects a plus sign in formulas (outside of `SUM(...)`, `AVERAGE(...)`, or `MAX(...)`), it flags it.

- **Accounting Number Format**  
  Checks if at least one cell in a typical income/expense range is set to Accounting format.

- **Offset Handling**  
  Attempts to handle an extra row if students inserted one after "Income," so it can still find key labels.

- **Bottom Borders**  
  Checks that certain rows (above `Total Income` and `Total Expenses`) have a bottom border.

- **Scoring**  
  Assigns a maximum of 70 points (excluding the chart portion). Deductions reduce the total from 70, and the script prints a final score as `X/100`.

---

## Usage

1. Download the script (for example, `grader.py`) and place your Excel files (`.xlsx`) in the same directory.
2. Run the script by typing: python grader.py
3. The script scans all `.xlsx` files in the directory, grades each one, and prints:
   - The final score (`X/100`, where 70 is the max ignoring charts)
   - Any deductions made

---

## Chart Limitation

The script does **not** handle chart grading because `openpyxl` cannot reliably read or interpret existing Excel charts from a saved file. You or your instructor will need to manually verify the chart portion.

---

## Reporting Issues

If you run into edge cases (like multiple sheets, more than one row offset, or unusual formatting) that the script doesnt handle, feel free to open an issue or reach out. Im happy to add fixes or features.

---
## Results 
After Running Python Autograde.py

![image](https://github.com/user-attachments/assets/de005eec-66a4-4671-bb28-de27f0351d83)

What the incorrect Sheet looks like 
![image](https://github.com/user-attachments/assets/4dccf2aa-3528-466b-bcbe-00776e975b20)

What a mosly correct sheet looks like 
![image](https://github.com/user-attachments/assets/51e9ddbc-ede1-4285-b204-1749fc8acdaa)



