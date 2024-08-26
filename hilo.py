
import pandas as pd
import numpy as np
import datetime

# Getting data from excel wb
file_path = '/Users/emilyyuldasheva/Documents/Практика 2024/prac.xlsx'
data = pd.read_excel(file_path)

# Combine Date and Time to DateTime
data['DateTime'] = data.apply(lambda row: datetime.datetime.combine(row['Date'], row['Time']), axis=1)

# Group the data by tags and sort each group by timestamp
grouped_data = data.groupby('Tag').apply(lambda x: x.sort_values(by='DateTime')).reset_index(drop=True)


results = []

# Check all the unique tags
for tag, group in grouped_data.groupby('Tag'):
    suppressed_alarms = 0
    data_array = group[['DateTime', 'Description.2']].to_numpy()
    
    lo_indices = np.where(data_array[:, 1] == 'ALM')[0]
    recover_indices = np.where(data_array[:, 1] == 'NR')[0]

    i = 0
    while i < len(lo_indices):
        lo_time = data_array[lo_indices[i], 0]
        pairs_in_window = 0

        for j in recover_indices:
            if j > lo_indices[i]:
                recover_time = data_array[j, 0]
                if (recover_time - lo_time).total_seconds() <= 300:
                    pairs_in_window += 1
                    lo_time = recover_time
                else:
                    break

        if pairs_in_window > 3:
            suppressed_alarms += pairs_in_window - 3  # Count only those after the very first three in the series

        i += 1
        while i < len(lo_indices) and lo_indices[i] <= j:
            i += 1

    total_alarms = len(lo_indices)
    results.append([tag, total_alarms, suppressed_alarms])

# Create DataFrame having the results obtained
results_df = pd.DataFrame(results, columns=['Tag', 'Total Alarms', 'Suppressed Alarms'])

# sum of all the suppressed alarms
total_suppressed_alarms = results_df['Suppressed Alarms'].sum()
total_alarms = results_df['Total Alarms'].sum()
results_df.loc[len(results_df)] = ['Total', total_alarms, total_suppressed_alarms]

# Save DayaFrame in Excel wb
output_file_path = file_path.replace('.xlsx', '_results.xlsx')
results_df.to_excel(output_file_path, index=False)


print(f"Результаты сохранены в файл: {output_file_path}")
