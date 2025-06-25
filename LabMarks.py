import pandas as pd

"""
Student Lab Marks Analyzer

This code automates the analysis of student lab results stored in an Excel file.

Features:
- Calculates percentage scores for multiple labs based on total marks
- Computes the average percentage for each student
- Assigns a final Pass/Fail status based on average
- Outputs a cleaned DataFrame showing only percentages, average, and status

This code saves time and optimizes accuracy when marking multiple chemistry labs for lots of students.
"""


# Load Excel file
df = pd.read_excel("Lab Marks.xlsx")


# The row 30 contains total marks
totals = df.iloc[30]
lab_columns = df.columns[1:]  # Skip the table title 'Student Number'


# Convert the marks to a percentage for each student and each lab mark
for lab in lab_columns:
    total_mark = totals[lab]
    percentages = (df.loc[:29, lab] / total_mark) * 100
    df.loc[:29, lab + ' %'] = percentages.round().astype('Int64')
    df.loc[30, lab + ' %'] = pd.NA  # Leave total row blank

# Get all percentage column names
percentage_columns = [lab + ' %' for lab in lab_columns]


# Calculate average % for each student for all lab marks. Row 30 is the total, not a student mark.
df.loc[:29, 'Average %'] = df.loc[:29, percentage_columns].mean(axis=1).round().astype('Int64')
df.loc[30, 'Average %'] = pd.NA


# Assign overall pass/fail status based on average
overall_status = []
for avg in df.loc[:29, 'Average %']:
    if 50 <= avg < 75:
        overall_status.append('Passed')
    elif avg >= 75:
        overall_status.append('Passed with Distinction')
    else:
        overall_status.append('Failed')

df.loc[:29, 'Overall Status'] = overall_status
df.loc[30, 'Overall Status'] = pd.NA #This avoids the last row from computing since it is showing the total marks that the lab was worth and not a student mark.


# Keep only student number, percentages, average, and overall status. Remove the original marks for simplicity.
columns_to_show = ['Student Number'] + percentage_columns + ['Average %', 'Overall Status']
df = df[columns_to_show]

# Show the final DataFrame
print(df)

# Save and export the new updated dataframe as an Excel file
df.to_excel('Analyzed Lab Marks.xlsx', index=False)



