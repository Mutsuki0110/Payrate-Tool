import pandas as pd

# Prompt the user to input the date
date = input("Please enter the date (e.g., 7.7.24): ")

# Format the file names using the input date
emec_file = f'nj1_plain_records/EMEC JDL NJ1 {date}.xlsx'
osi_file = f'nj1_plain_records/OSI JDL NJ1 {date}.xlsx'
wfe_file = f'nj1_plain_records/WFE JDL NJ1 {date}.xlsx'

# Read the Excel files
emec = pd.read_excel(emec_file)
osi = pd.read_excel(osi_file)
wfe = pd.read_excel(wfe_file)

# Combine the dataframes based on the pattern of the original file
combined_df = pd.concat([emec, osi, wfe], ignore_index=True)

# Rename the first column to "Workers' Full Name"
combined_df.rename(columns={combined_df.columns[0]: "Workers' Full Name"}, inplace=True)

# Ensure all values in "Workers' Full Name" are strings
combined_df["Workers' Full Name"] = combined_df["Workers' Full Name"].astype(str)

# Create "Workers' Last Name" and "Workers' First Name" columns
combined_df["Workers' Last Name"] = combined_df["Workers' Full Name"].apply(lambda x: x.split(' ', 1)[-1] if ' ' in x else '')
combined_df["Workers' First Name"] = combined_df["Workers' Full Name"].apply(lambda x: x.split(' ', 1)[0] if ' ' in x else x)

# Reorder columns to place "Workers' Last Name" and "Workers' First Name" after "Workers' Full Name"
cols = list(combined_df.columns)
cols.insert(1, cols.pop(cols.index("Workers' Last Name")))
cols.insert(2, cols.pop(cols.index("Workers' First Name")))
combined_df = combined_df[cols]

# Add a new column "Building" with the value "NJ1"
combined_df.insert(0, 'Building', 'NJ1')

# Insert a new column "Agency" in the second place with empty values
combined_df.insert(1, 'Agency', '')

# Rename the column "Grand Total" to "Weekly Total"
combined_df.rename(columns={"Grand Total": "Weekly Total"}, inplace=True)

# Keep only the first 13 columns
combined_df = combined_df.iloc[:, :13]

# Save the combined dataframe to a new Excel file
output_file = f'outputs/NJ1 {date}.xlsx'
combined_df.to_excel(output_file, index=False)
