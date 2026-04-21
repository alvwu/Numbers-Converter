# GradeBook Transfer

A Streamlit web application that converts Apple Numbers gradebook files into organized Excel workbooks with individual student sheets. Includes both grade transfer and attendance tracking features.

## Features

### Grade Transfer Tab
- **Numbers to Excel Conversion**: Upload `.numbers` files and convert them to `.xlsx` format
- **Individual Student Sheets**: Each student gets their own Excel sheet named "Last Name, First Name"
- **Student Identification**: Separate columns for ID, First Name, and Last Name displayed at the top of each sheet
- **Category Organization**: Automatically categorize grades by keywords (Exams, Assignments, Participation, El Civics, etc.)
- **Configurable Max Points**: Set the maximum points for each category (e.g., Exams out of 100, Participation out of 1)
- **Weighted Grading**: Assign weight percentages to each category for final grade calculation
- **Empty Row Filtering**: Automatically skips rows with blank student identifiers

### Attendance Tab
- **Attendance Tracking**: Convert attendance data from Numbers to organized Excel sheets
- **Individual Student Sheets**: Each student gets their own sheet with attendance records
- **Flexible Student Columns**: Attendance files can use any combination of ID, First Name, and Last Name columns
- **Auto-Detected Dates**: Date-like attendance columns are automatically detected and selected
- **Chronological Export**: Attendance dates are written to Excel in date order
- **Color-Coded Grades**:
  - Green highlight for present (1)
  - Orange highlight for absent (0)
- **Attendance Summary**: Each sheet includes days present, total days, and attendance rate percentage

## Installation

1. Clone or download this repository

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   streamlit run app.py
   ```

## Usage

### Grade Transfer
1. **Upload**: Drag and drop your `.numbers` gradebook file in the "Grade Transfer" tab
2. **Select Columns**: Choose which columns contain Student ID, First Name, and Last Name
3. **Configure Categories** (in sidebar):
   - Add/remove categories with keywords
   - Set max points per item for each category
   - Set weight percentage for each category (should total 100%)
4. **Generate**: Click "Generate Excel File" to create the workbook
5. **Download**: Download the organized Excel file

### Attendance
1. **Upload**: Drag and drop your `.numbers` attendance file in the "Attendance" tab
2. **Select Columns**: Choose whichever student identifier columns are available (ID, First Name, and/or Last Name)
3. **Select Dates**: Review the auto-detected attendance date columns and adjust them if needed
4. **Generate**: Click "Generate Attendance Excel" to create the workbook
5. **Download**: Download the color-coded attendance Excel file

## Excel Output Structure

### Grade Transfer Output
Each student sheet contains:

| Section | Description |
|---------|-------------|
| **Header** | ID, First Name, Last Name |
| **Grades by Category** | Assignment names, scores, and max points organized by category |
| **Category Averages** | (Optional) Percentage average for each category |
| **Weighted Grades** | Breakdown showing each category's score, weight, and weighted contribution |
| **Final Weighted Grade** | The calculated final grade based on category weights |

### Attendance Output
Each student sheet contains:

| Section | Description |
|---------|-------------|
| **Header** | Any available student identifiers (ID, First Name, Last Name) |
| **Attendance Records** | Date and attendance value (0 or 1), color-coded and ordered chronologically |
| **Summary** | Days present, total days, and attendance rate percentage |

## Default Categories

| Category | Keywords | Default Max Points | Default Weight |
|----------|----------|-------------------|----------------|
| Exams | exam, test, midterm, final | 100 | 25% |
| Assignments | assignment, homework, hw | 100 | 25% |
| Participation | participation, attendance | 1 | 30% |
| El Civics | el civics, civics, elcivics | 100 | 20% |
| Other | (uncategorized items) | 100 | 0% |

## Requirements

- Python 3.7+
- streamlit
- pandas
- openpyxl
- numbers-parser

## License

Made for educators.
