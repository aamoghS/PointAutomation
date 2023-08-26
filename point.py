import pandas as pd
from datetime import datetime

#First Excel file
file1_data = pd.read_excel('file1.xlsx')

#Second Excel file
file2_data = pd.read_excel('file2.xlsx')

# Group names
points_dict = file1_data.groupby(['First Name', 'Last Name'])['Points'].sum().to_dict()

# merage the data
merged_data = []
for index, row in file2_data.iterrows():
    first_name = row['First Name']
    last_name = row['Last Name']
    points = points_dict.get((first_name, last_name), 0) + 1
    prev_points = points_dict.get((first_name, last_name), 0)  # Previous points
    points_difference = points - prev_points  # Calculate points difference
    date_added = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Get current date and time
    merged_data.append({
        'First Name': first_name,
        'Last Name': last_name,
        'Points': points,
        'Points Added': points_difference,
        'Date Added': date_added
    })

# to df 
merged_df = pd.DataFrame(merged_data)

# save as excel
merged_df.to_excel('merged_file.xlsx', index=False)

print("Excel files merged and updated.")
