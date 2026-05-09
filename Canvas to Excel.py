import pandas as pd
from ics import Calendar
import requests
from io import StringIO
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import re

# Load ICS Calendar 
url = input("Enter Canvas calendar ICS feed URL or path to local .ics file: ")
if url.startswith('http'):
    r = requests.get(url)
    c = Calendar(r.text)
else:
    with open(url, 'r', encoding='utf-8') as f:
        c = Calendar(f.read())

# Parse all events from ICS 
data = []
for event in c.events:
    if event.begin:
        event_date = event.begin.datetime
        assignment = event.name.strip()

        # Try to extract course from square brackets
        match = re.search(r'\[(.*?)\]', assignment)
        if match:
            course = match.group(1).strip()
            assignment = assignment.replace(f'[{course}]', '').strip()
        else:
            course = "Unknown"

        data.append({
            "Date Group": event_date.date(),
            "Assignment": assignment,
            "Course": course,
            "Due Date": event_date.date(),
            "Notes": ""
        })

# Organize into DataFrame 
df = pd.DataFrame(data)

# Sort by date then course then assignment
df = df.sort_values(by=["Date Group", "Course", "Assignment"])

# Assign Colors for Courses 
unique_courses = [course for course in df['Course'].unique() if course != "Unknown"]
color_palette = [
    "FFFF99", "99CCFF", "FFCC99", "CCFFCC", "FF9999", "CCCCFF", "FFB6C1", "D3FFCE", "FFD700", "ADFF2F",
    "E6E6FA", "E0FFFF", "F0FFF0", "FFFACD", "F5DEB3", "EED5D2", "D8BFD8", "B0E0E6", "FAFAD2", "FFE4E1"
]  # Extended pastel palette

course_colors = {course: color_palette[i % len(color_palette)] for i, course in enumerate(unique_courses)}

# Export to Excel with Formatting
wb = Workbook()
ws = wb.active
ws.title = "Assignments"

# Write headers
headers = ["Assignment", "Course", "Due Date", "Notes"]
for col_num, header in enumerate(headers, 1):
    ws.cell(row=1, column=col_num, value=header)

current_day = None
row_num = 2

for _, row in df.iterrows():
    if current_day != row['Date Group']:
        if current_day is not None:
            row_num += 1  # blank row between days
        current_day = row['Date Group']

    ws.cell(row=row_num, column=1, value=row['Assignment'])
    ws.cell(row=row_num, column=2, value=row['Course'])
    ws.cell(row=row_num, column=3, value=str(row['Due Date']))
    ws.cell(row=row_num, column=4, value=row['Notes'])

    # Color the Course cell (skip 'Unknown')
    if row['Course'] != "Unknown":
        fill = PatternFill(start_color=course_colors[row['Course']], end_color=course_colors[row['Course']], fill_type="solid")
        ws.cell(row=row_num, column=2).fill = fill

    row_num += 1

# Autosize columns
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width

# Make Notes column specifically 500 pixels wide (about 70 width units in openpyxl)
ws.column_dimensions[get_column_letter(4)].width = 70

# Save
wb.save("canvas_assignments.xlsx")

print("Done! Output saved to canvas_assignments.xlsx")
