import pandas as pd

# First Excel file
file1_data = pd.read_excel('file1.xlsx')

# Second Excel file
file2_data = pd.read_excel('file2.xlsx')

# group the files 
points_dict = file1_data.groupby(['First Name', 'Last Name'])['Points'].sum().to_dict()

# Merge mile and add points 
merged_data = []
for index, row in file2_data.iterrows():
    first_name = row['First Name']
    last_name = row['Last Name']
    points = points_dict.get((first_name, last_name), 0) + 1
    merged_data.append({'First Name': first_name, 'Last Name': last_name, 'Points': points})

# merged to a frame 
merged_df = pd.DataFrame(merged_data)

# additional
additional_columns = list(file2_data.columns)
additional_columns.remove('First Name')
additional_columns.remove('Last Name')

for col in additional_columns:
    merged_df[col] = file2_data[col]

# saves
merged_df.to_excel('merged_file.xlsx', index=False)

print("Updated.")
