import pandas as pd
from datetime import timedelta
from formatting import apply_formatting
import openpyxl

def extract_and_transform(filename, sheet_name, filter_value):
    xl = pd.ExcelFile(filename)
    df = xl.parse(sheet_name)

    df_filtered = df[df['Labels'] == filter_value].copy()

    df_filtered['Created'] = pd.to_datetime(df_filtered['Created']).dt.date
    df_filtered['Resolved'] = pd.to_datetime(df_filtered['Resolved']).dt.date

    # Count unique dates from 'new defects' column
    new_defects_counts = df_filtered['Created'].value_counts().reset_index()
    new_defects_counts.columns = ['Date', 'Created - Daily']

    # Count unique dates from 'closed defects' column
    closed_defects_counts = df_filtered['Resolved'].value_counts().reset_index()
    closed_defects_counts.columns = ['Date', 'Closed - Daily']

    # Set start_date from 'Created' column since it's always available
    start_date = new_defects_counts['Date'].min()

# For end_date, check if 'Resolved' has any non-NaT dates.
# If so, pick the latest date from both columns; otherwise, rely only on 'Created'.
    if closed_defects_counts['Date'].dropna().empty:
        end_date = new_defects_counts['Date'].max()
    else:
        end_date = max(new_defects_counts['Date'].max(), closed_defects_counts['Date'].dropna().max())

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

    return df_continuous

def write_to_excel(filename, dataframe, sheet_name):
    # Open the Excel file using openpyxl
    book = openpyxl.load_workbook(filename)

    # Check if the sheet exists, and if so, delete it
    if sheet_name in book.sheetnames:
        del book[sheet_name]

    # Save the changes and close the workbook
    book.save(filename)
    book.close()

    # Write to new sheet
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
        dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]

        # Apply the formatting
        apply_formatting(worksheet)

file_name = "C:\\Users\\foster.nilsson\\Downloads\\Jira (3).xlsx"
sheet_name = "Overall_Defect Jira Exports"

# Process for "W9_BenefitsReporting"
df_br_cumulative = extract_and_transform(file_name, sheet_name, "W9_BenefitsReporting")
write_to_excel(file_name, df_br_cumulative, "BR Cumulative")

# Process for "Mock_1_CO"
df_mock_cumulative = extract_and_transform(file_name, sheet_name, "Mock_1_CO")
write_to_excel(file_name, df_mock_cumulative, "Mock Cumulative")   