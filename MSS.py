import pandas as pd
import re
import glob
from pathlib import Path
import os

file = 'Ready MSS Report.xlsx'
# link_pattern = r"^[A-Za-z]{2}\d{4}[-_][A-Za-z]\d{4}$"

home_directory = os.path.expanduser('~')
csv_files = glob.glob(os.path.join(home_directory, '**', '*.csv'), recursive=True)
csv_files.sort(key=os.path.getctime, reverse=True)
csv_file_path = csv_files[0]

mss_df = pd.read_csv(csv_file_path, sep=',') # separatin' csv file
link_df = pd.read_excel('Alcatel Link Name.xlsx', sheet_name='RSL') # readin' file for findin' matches link name
mss_df.to_excel(file, index=False) # convert to excel for better user experience

mss_df = mss_df.drop( # droppin' useless columns for bad rx report
    columns=[
        'Time Logged', 'Elapsed Time', 'Elapsed Time Periodic', 'Period End Time', 'Period End Time Periodic', 'Suspect Interval Flag', 'Average Level Periodic (dBm)',
        'Granularity Period', 'Granularity Period Periodic', 'Maximum Level (dBm)', 'Maximum Level Periodic (dBm)', 'Minimum Level (dBm)', 'Minimum Level Periodic (dBm)',
        'Num Suppressed Intervals', 'Num Suppressed Intervals Periodic', 'Design vs Actual Deviation (dB)', 'Design vs Actual Deviation Periodic (dB)', 'Install vs Actual Deviation (dB)',
        'Install vs Actual Deviation Periodic (dB)', 'History Created','Periodic Time', 'Record Type','Suspect'
    ], axis=1
)
with pd.ExcelWriter(file, engine='openpyxl') as writer:
    mss_df.to_excel(writer, index=False)

mss_df['Time Captured'] = pd.to_datetime(mss_df['Time Captured'].str.split(' ').str[0],errors='coerce').dt.date # reformattin this Series

mss_df = mss_df.pivot_table(
    index='Monitored Object',
    columns='Time Captured',
    values='Average Level (dBm)',
    aggfunc='mean'
)
mss_df['Mean RSL'] = mss_df.mean(axis=1)
mss_df = mss_df[~mss_df['Mean RSL'].between(-100, -80)]
mss_df = mss_df[~mss_df['Mean RSL'].between(-48, 0)]

mss_df = pd.merge(
    mss_df, link_df[['Monitored Object', 'link name']], on='Monitored Object', how='inner'
).reset_index(drop=True)

mss_df['link name'] = mss_df['link name'].str.slice(0,13)
mss_df['far end'] = mss_df['link name'].apply(lambda x: '-'.join(x.split('-')[::-1]) if '-' in x else '_'.join(x.split('_')[::-1]))

with pd.ExcelWriter(file, engine='openpyxl', mode='a') as writer:
    mss_df.to_excel(writer, sheet_name='RSL PIVOT', index=False)

# mss_df = pd.merge(
#     mss_df, link_df[['link name']], how='inner', left_on='Monitored Object', right_on='link name' 
# ).reset_index()
# print(mss_df)
# with pd.ExcelWriter(file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
#     mss_df.to_excel(writer, sheet_name='RSL PIVOT')