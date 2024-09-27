# import pandas as pd
# import numpy as np
# import glob
# import os
# from pathlib import Path
# import re

# NR_REPORT = "Ready NR Report.xlsx"

# home_directory = os.path.expanduser('~')  # user's home directory
# # все файлы с расширением .xlsx в домашней директории и всех её подкаталогах
# files = glob.glob(os.path.join(home_directory, '**', '*.xlsx'), recursive=True)
# # сортируем файлы по времени создания в порядке убывания (последние созданные/добавленные файлы будут первыми)
# files.sort(key=os.path.getctime, reverse=True)
# latest_files = files[:3]  # три последних добавленных файла
# dataframes = [pd.read_excel(file, skiprows=range(0, 5), sheet_name='sheet1')
#               # read files into DF and skipped unusable rows
#               for file in latest_files]

# for path in latest_files:
#     print(Path(path))
# # определяю переменные на основе имени файла
# new_nr, old_nr1, old_nr2 = None, None, None
# for file, df in zip(latest_files, dataframes):
#     if "NR8120" in file:
#         old_nr1 = df
#     elif "NR8250" in file:
#         old_nr2 = df
#     elif "checkpoint" in file:
#         new_nr = df

# if any(var is None for var in (new_nr, old_nr1, old_nr2)):
#     print("Unable to find all required files matching the given criteria.")
# else:
#     print("All files were successfully read into the df and linked to variables.")

# columns_to_drop_new = ['Index', 'End Time', 'Query Granularity',
#                        'Neighbor NE Ip', 'Neighbor NE Port', 'IPADDRESS', 'LINK NAME']

# columns_to_drop_old = ['Index', 'End Time', 'Query Granularity',
#                        'IP Address', 'Neighbor NE IP', 'Neighbor NE Port', 'LINK NAME']
# new_nr = new_nr.drop(
#     columns=[col for col in columns_to_drop_new if col in new_nr.columns], axis=1)
# old_nr1 = old_nr1.drop(
#     columns=[col for col in columns_to_drop_old if col in old_nr1.columns], axis=1)
# old_nr2 = old_nr2.drop(
#     columns=[col for col in columns_to_drop_old if col in old_nr2.columns], axis=1)

# # Объединение всех DataFrame
# all_nr = pd.concat([new_nr, old_nr1, old_nr2], ignore_index=True)
# all_nr['Start Time'] = pd.to_datetime(
#     all_nr['Start Time']).dt.date  # Преобразование даты и времени
# all_nr['Link Name'] = all_nr['NE Location'].str.split(',').str[-1]

# with pd.ExcelWriter(NR_REPORT, engine='openpyxl') as writer:
#     all_nr.to_excel(writer, sheet_name='Test Page', index=False)

# all_nr_agg = all_nr[['Link Name', 'MO Location', 'Mean Transmitted Power(dBm)', 'Mean Received Signal Level(dBm)']].groupby(['Link Name', 'MO Location']).agg(
#     tsl_count=pd.NamedAgg(
#         column='Mean Transmitted Power(dBm)', aggfunc='mean'),
#     rsl_count=pd.NamedAgg(
#         column='Mean Received Signal Level(dBm)', aggfunc='mean')
# ).reset_index()

# all_nr_agg = all_nr_agg[
#     ~all_nr_agg['tsl_count'].between(-100, 0)
# ]
# all_nr_agg = all_nr_agg[
#     ~all_nr_agg['rsl_count'].between(-100, -80)
# ]
# all_nr_agg = all_nr_agg[
#     ~all_nr_agg['rsl_count'].between(-48, 0)
# ]

# all_nr_agg = all_nr_agg.reset_index()


# all_nr_agg['Site List'] = [
#     [] for _ in range(len(all_nr_agg))
# ]

# all_nr_agg['LINK'] = pd.Series(dtype='object')


# regex = re.compile(r'^[A-Za-z]{2}\d{4}$|^[A-Za-z]{3}\d{3}$|^[A-Za-z]{4}\d{2}')

# for i in all_nr_agg.index:
#     print(i)
#     all_nr_agg['Site List'][i] = re.split(
#         r'[-_. ]', all_nr_agg['Link Name'][i])
#     all_nr_agg['Site List'][i] = [
#         j for j in all_nr_agg['Site List'][i] if regex.search(j)]

# all_nr_agg['Site List Str'] = all_nr_agg['Site List'].apply(
#     lambda x: ','.join(x))

# min_rsl_idx = all_nr_agg.groupby('Site List Str')['rsl_count'].idxmin()
# all_nr_agg = all_nr_agg.loc[min_rsl_idx]

# all_nr_agg = all_nr_agg.drop(columns=[
#     'Site List Str', 'tsl_count'
# ])
# all_nr_agg = all_nr_agg.reset_index()


# def check_two_array(list1, list2):
#     return list1[0] in list2[1:] and list2[0] in list1[1:]


# for i in all_nr_agg.index:
#     if not pd.isna(all_nr_agg['LINK'][i]):
#         continue
#     link_name = all_nr_agg['Link Name'][i]
#     site_lst = all_nr_agg['Site List'][i]
#     all_nr_agg['LINK'][i] = link_name
#     for j in range(i+1, len(all_nr_agg)):
#         if not pd.isna(all_nr_agg['LINK'][j]):
#             continue
#         if check_two_array(site_lst, all_nr_agg['Site List'][j]):
#             all_nr_agg['LINK'][j] = link_name
#         else:
#             continue

# all_nr_agg = all_nr_agg.drop(columns=[
#     'level_0', 'index'
# ])

# rsl_min_idx = all_nr_agg.groupby('LINK')['rsl_count'].idxmin()
# all_nr_agg = all_nr_agg.loc[rsl_min_idx]

# with pd.ExcelWriter(NR_REPORT, engine='openpyxl', mode='a') as writer:
#     all_nr_agg.to_excel(writer, sheet_name='Test Page 2', index=False)


import pandas as pd
import numpy as np
import glob
import os
from pathlib import Path

NR_REPORT = "Ready NR Report.xlsx"

home_directory = os.path.expanduser('~')  # user's home directory
# все файлы с расширением .xlsx в домашней директории и всех её подкаталогах
files = glob.glob(os.path.join(home_directory, '**', '*.xlsx'), recursive=True)
# сортируем файлы по времени создания в порядке убывания (последние созданные/добавленные файлы будут первыми)
files.sort(key=os.path.getctime, reverse=True)
latest_files = files[:3]  # три последних добавленных файла
dataframes = [pd.read_excel(file, skiprows=range(0, 5), sheet_name='sheet1')
              # read files into DF and skipped unusable rows
              for file in latest_files]

for path in latest_files:
    print(Path(path))
# определяю переменные на основе имени файла
new_nr, old_nr1, old_nr2 = None, None, None
for file, df in zip(latest_files, dataframes):
    if "NR8120" in file:
        old_nr1 = df
    elif "NR8250" in file:
        old_nr2 = df
    elif "checkpoint" in file:
        new_nr = df

if any(var is None for var in (new_nr, old_nr1, old_nr2)):
    print("Unable to find all required files matching the given criteria.")
else:
    print("All files were successfully read into the df and linked to variables.")

columns_to_drop_new = ['Index', 'End Time', 'Query Granularity',
                       'Neighbor NE Ip', 'Neighbor NE Port', 'IPADDRESS', 'LINK NAME']
columns_to_drop_old = ['Index', 'End Time', 'Query Granularity',
                       'IP Address', 'Neighbor NE IP', 'Neighbor NE Port', 'LINK NAME']
new_nr = new_nr.drop(
    columns=[col for col in columns_to_drop_new if col in new_nr.columns], axis=1)
old_nr1 = old_nr1.drop(
    columns=[col for col in columns_to_drop_old if col in old_nr1.columns], axis=1)
old_nr2 = old_nr2.drop(
    columns=[col for col in columns_to_drop_old if col in old_nr2.columns], axis=1)

# Объединение всех DataFrame
all_nr = pd.concat([new_nr, old_nr1, old_nr2], ignore_index=True)

all_nr['Start Time'] = pd.to_datetime(
    all_nr['Start Time']).dt.date  # Преобразование даты и времени

# Создание столбца Full Name
all_nr['Full Name'] = all_nr['NE Location'].str.split(
    ',').str[-1] + '-' + all_nr['MO Location']
all_nr = all_nr.drop(columns=['NE Location', 'MO Location'], axis=1)
# Перемещение столбца Full Name
loc_index = all_nr.columns.get_loc('Neighbor NE Name') + 1
all_nr.insert(loc=loc_index, column='Full Name', value=all_nr.pop('Full Name'))
# Преобразование столбца RSL
all_nr['Mean Received Signal Level(dBm)'] = pd.to_numeric(
    all_nr['Mean Received Signal Level(dBm)'].astype(str).str.replace(',', '.'), errors='coerce'
)
all_nr.to_excel(NR_REPORT, index=False)

# Создание pivot table для TSL и RSL
tsl_table = all_nr.pivot_table(index='Full Name', columns='Start Time',
                               values='Mean Transmitted Power(dBm)', aggfunc='mean')
tsl_table['Mean TSL'] = tsl_table.mean(axis=1)

rsl_table = all_nr.pivot_table(index='Full Name', columns='Start Time',
                               values='Mean Received Signal Level(dBm)', aggfunc='mean')
rsl_table['Mean RSL'] = rsl_table.mean(axis=1)

# Создаем новые страницы и сохраняем pivot table в них
with pd.ExcelWriter(NR_REPORT, engine='openpyxl', mode='a') as writer:
    tsl_table.to_excel(writer, sheet_name='TSL PIVOT')
    rsl_table.to_excel(writer, sheet_name='RSL PIVOT')

all_nr = pd.merge(tsl_table[['Mean TSL']], rsl_table,
                  on='Full Name', how='inner').reset_index()

# if 'Full Name' not in inner_join.columns:
#     inner_join = inner_join.reset_index()

# inner_join = inner_join[[col for col in inner_join.columns if col != 'Mean TSL'] + ['Mean TSL']]
all_nr = all_nr[~all_nr['Mean TSL'].between(-100, 0)]
all_nr = all_nr[~all_nr['Mean RSL'].between(-100, -80)]
all_nr = all_nr[~all_nr['Mean RSL'].between(-48, 0)]
all_nr = all_nr.drop(columns='Mean TSL', axis=1)

columns_to_check = all_nr.columns.difference(
    ['Full Name', 'Mean RSL', 'Neighbor NE Name'])
all_nr[columns_to_check] = all_nr[columns_to_check].fillna(-9999)
mask = (all_nr[columns_to_check] > -48).sum(axis=1) >= 2
all_nr = all_nr[~mask]
all_nr[columns_to_check] = all_nr[columns_to_check].replace(-9999, np.nan)

# Чтение и объединение с другим листом Excel
sheet1 = pd.read_excel(NR_REPORT, sheet_name='Sheet1')
all_nr = pd.merge(all_nr, sheet1[['Full Name', 'Neighbor NE Name']],
                  on='Full Name', how='left').drop_duplicates(subset=['Full Name'])

all_nr['Neighbor NE Name'] = all_nr['Neighbor NE Name'].str.strip()
all_nr['Neighbor NE Name'].replace('', np.nan)
all_nr['Ready Name[A-B]'] = all_nr.apply(
    lambda row: (row['Full Name'][:7] + row['Neighbor NE Name']
                 [:6]) if pd.notna(row['Neighbor NE Name']) else '',
    axis=1
)
all_nr['Reversed Name[B-A]'] = all_nr.apply(
    lambda row: (row['Neighbor NE Name'][:7] + row['Ready Name[A-B]']
                 [:6]) if pd.notna(row['Neighbor NE Name']) else '',
    axis=1
)

with pd.ExcelWriter(NR_REPORT, engine='openpyxl', mode='a') as writer:
    all_nr.to_excel(writer, sheet_name='RSL & TSL MEAN', index=False)
