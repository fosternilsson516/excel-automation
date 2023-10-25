import pandas as pd
from datetime import timedelta
from formatting import apply_formatting


# Load the Excel file
file_name = "C:\\Users\\foster.nilsson\\Downloads\\AMIW9 Defect Count Dashboard.xlsx"
xl = pd.ExcelFile(file_name)

# Select the specific sheet
sheet_name = "Overall_Defect Jira Exports"
df = xl.parse(sheet_name)

# Filter data based on a specific project group
filtered_group = "W9_BenefitsReporting"
df_filtered = df[df['Labels'] == filtered_group].copy()

df_filtered['Created'] = pd.to_datetime(df_filtered['Created']).dt.date
df_filtered['Resolved'] = pd.to_datetime(df_filtered['Resolved']).dt.date

# Count unique dates from 'new defects' column
new_defects_counts = df_filtered['Created'].value_counts().reset_index()
new_defects_counts.columns = ['Date', 'Created - Daily']

# Count unique dates from 'closed defects' column
closed_defects_counts = df_filtered['Resolved'].value_counts().reset_index()
closed_defects_counts.columns = ['Date', 'Closed - Daily']

start_date = min(new_defects_counts['Date'].min(), closed_defects_counts['Date'].min())
end_date = max(new_defects_counts['Date'].max(), closed_defects_counts['Date'].max())
all_dates = [start_date + timedelta(days=x) for x in range(0, (end_date-start_date).days + 1)]
df_continuous = pd.DataFrame(all_dates, columns=['Date'])

# Merge this DataFrame with the counts DataFrames
df_continuous = pd.merge(df_continuous, new_defects_counts, on='Date', how='left')
df_continuous = pd.merge(df_continuous, closed_defects_counts, on='Date', how='left')

# Fill NaN with zeros
df_continuous.fillna(0, inplace=True)
df_continuous['Created - Daily'] = df_continuous['Created - Daily'].astype(int)
df_continuous['Created - Cumulative'] = df_continuous['Created - Daily'].cumsum()
df_continuous['Closed - Daily'] = df_continuous['Closed - Daily'].astype(int)
df_continuous['Closed - Cumulative'] = df_continuous['Closed - Daily'].cumsum()

df_continuous = df_continuous[['Date', 'Created - Daily', 'Created - Cumulative', 'Closed - Daily', 'Closed - Cumulative']]

#print(df_continuous)

import openpyxl

# Open the Excel file using openpyxl
book = openpyxl.load_workbook(file_name)

# Check if the sheet exists, and if so, delete it
if "BR Cumulative2" in book.sheetnames:
    del book["BR Cumulative2"]

# Save the changes and close the workbook
book.save(file_name)
book.close()

# Write to new sheet
with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
    df_continuous.to_excel(writer, sheet_name="BR Cumulative2", index=False)
    worksheet = writer.sheets["BR Cumulative2"]

    # Apply the formatting
    apply_formatting(worksheet)
   