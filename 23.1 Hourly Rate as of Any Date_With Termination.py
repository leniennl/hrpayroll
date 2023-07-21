import pandas as pd

# Change this date as wish
SPECIFIED_DATE = pd.to_datetime('2022-07-04')

# Read the original spreadsheet with rate changing hisotry, exported from Payforce
HISTORY_DF = pd.read_csv(r"H:\PAYRATE17.07.CSV")

# Convert dates to pandas datetime objects
HISTORY_DF['CHANGE DATE'] = pd.to_datetime(HISTORY_DF['CHANGE DATE'])
HISTORY_DF['EFFECTIVE_DATE'] = pd.to_datetime(HISTORY_DF['EFFECTIVE_DATE'])
HISTORY_DF['Term Date'] = pd.to_datetime(HISTORY_DF['Term Date'])

# fill effective date with change date:
HISTORY_DF['EFFECTIVE_DATE'].fillna(HISTORY_DF['CHANGE DATE'], inplace=True)

# Fill null values in 'Term Date' column with a future date
HISTORY_DF['Term Date'].fillna(pd.Timestamp.max, inplace=True)

# filter the df for effective date smaller than the specified date:
filtered_df = HISTORY_DF[(HISTORY_DF['EFFECTIVE_DATE'] <= SPECIFIED_DATE) 
                         & (HISTORY_DF['Term Date'] >= SPECIFIED_DATE )]
filtered_df= filtered_df.drop_duplicates()

# Replace pd.Timestamp.max with null
filtered_df.replace(pd.Timestamp.max, pd.NaT, inplace=True)
filtered_df.reset_index(drop=True, inplace=True)

# Group by each Employee code and Employee Name combination
grouped_df = filtered_df.groupby(['Employee code'])

# Sort by Effective_date from the latest to older dates and Change Date from the latest to earlier dates
sorted_df = grouped_df.apply(lambda x: x.sort_values(['EFFECTIVE_DATE', 'CHANGE DATE'], ascending=[False, False]))
sorted_df.reset_index(drop=True, inplace=True)

# Select the top row for each group
new_df = sorted_df.groupby('Employee code').head(1)

# Keep only the desired columns
columns_to_keep = ['Employee code', 'Employee Name', 'CHANGE DATE', 'EFFECTIVE_DATE', 'Amount','Term Date']
new_df = new_df.loc[:, columns_to_keep]

# Format 'CHANGE DATE' column to display only the date in 'dd/mm/yyyy' format
new_df['CHANGE DATE'] = new_df['CHANGE DATE'].dt.strftime('%d/%m/%Y')
new_df['EFFECTIVE_DATE'] = new_df['EFFECTIVE_DATE'].dt.strftime('%d/%m/%Y')
new_df['Term Date'] = new_df['Term Date'].dt.strftime('%d/%m/%Y')

# Remove rows where 'Employee Name' is missing or empty
new_df = new_df.dropna(subset=['Employee Name'])

# create a new excel file
new_df.to_excel('h:/Hourly Rate as at ' + SPECIFIED_DATE.strftime('%d %b %Y')+ '.xlsx', index=False)

print("File Generated at H:\ root drive.")