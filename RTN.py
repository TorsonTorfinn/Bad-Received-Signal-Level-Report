import pandas as pd
import numpy as np
import os
import re
import zipfile
from datetime import datetime
from openpyxl import load_workbook
import win32com.client as win32

directory_path = r"C:/Users/suhrob.yusubaxmedov/downloads"
dest_path = r"C:\projects\MSS_REPORT"

input_file_path = r'C:\projects\MSS_REPORT\RTN.xlsx'
output_file_path = r'C:\projects\MSS_REPORT\RTN_resaved.xlsx'
rtn_repot_path = r'C:\projects\MSS_REPORT\RTN_resaved.xlsx'

def extractor():
    # Регулярное выражение для поиска архивов с изменяющейся датой в названии
    pattern_zip = re.compile(r'History_Performance_Data_(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})\.zip')

    zip_files_with_dates = []  # Список для найденных архивов и их дат

    # Поиск всех архивов в указанной директории
    for file in os.listdir(directory_path):
        match = pattern_zip.match(file)
        if match:
            date_str = match.group(1)  # Парсинг даты из имени файла
            date_obj = datetime.strptime(date_str, '%Y-%m-%d_%H-%M-%S')
            zip_files_with_dates.append((file, date_obj))

    # Если найдены архивы, сортируем их по дате
    if zip_files_with_dates:
        zip_files_with_dates.sort(key=lambda x: x[1], reverse=True)

        latest_zip_file = zip_files_with_dates[0][0]  # Самый свежий и дикий Performance Data из загрузок :)
        zip_path = os.path.join(directory_path, latest_zip_file)

        # Распаковка архива
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(dest_path)

        # Поиск извлеченного файла
        extracted_rtn_file = os.listdir(dest_path)
        pattern_rtn_file = re.compile(r'History_Performance_Data_(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})\.xlsx')

        for file in extracted_rtn_file:
            if pattern_rtn_file.match(file):
                old_file_path = os.path.join(dest_path, file)
                new_file_path = os.path.join(dest_path, 'RTN.xlsx')  # Новое имя для файла History Performance Data
                if os.path.exists(new_file_path):
                    os.remove(new_file_path)
                os.rename(old_file_path, new_file_path)


def resave_excel_using_win32(input_path, output_path):
    try:
        excel = win32.Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(input_path)
        wb.SaveAs(output_path, FileFormat=51)  # FileFormat=51 означает сохранить как .xlsx
        wb.Close()
        excel.Application.Quit()
        print(f"Файл успешно пересохранен как: {output_path}")

        # Удаление исходного файла после пересохранения
        if os.path.exists(input_path):
            os.remove(input_path)
            print(f"Файл {input_path} удален.")
    except Exception as e:
        print(f"Ошибка при пересохранении файла с помощью Excel: {e}")


# Вызов функций
extractor()
resave_excel_using_win32(input_file_path, output_file_path)

rtn_df = pd.read_excel(rtn_repot_path, skiprows=7, sheet_name='Sheet1')

rtn_df['End Time'] = pd.to_datetime(rtn_df['End Time']).dt.date
tsl_df = rtn_df[rtn_df['Performance Event'] == 'TSL_AVG(dbm)']
rsl_df = rtn_df[rtn_df['Performance Event'] == 'RSL_AVG(dbm)']

tsl_pivot = tsl_df.pivot_table(
    index='Monitored Object', columns='End Time', values='Value CUR', aggfunc='mean'
)
tsl_pivot['Mean TSL'] = tsl_pivot.mean(axis=1)

rsl_pivot = rsl_df.pivot_table(
    index='Monitored Object', columns='End Time', values='Value CUR', aggfunc='mean'
)
rsl_pivot['Mean RSL'] = rsl_pivot.mean(axis=1)

# Объединяем данные по 'Monitored Object'
merged_df = pd.merge(tsl_pivot['Mean TSL'], rsl_pivot, on='Monitored Object', how='inner').reset_index()

# Применяем фильтры
merged_df = merged_df[~merged_df['Mean TSL'].between(-100, 0)]
merged_df = merged_df[~merged_df['Mean RSL'].between(-100, -80)]
merged_df = merged_df[~merged_df['Mean RSL'].between(-48, 0)]

# Определяем столбцы для проверки и обрабатываем NaN значения
rows_to_check = merged_df.columns.difference(['Monitored Object', 'Mean RSL'])
merged_df[rows_to_check] = merged_df[rows_to_check].fillna(-9999)

# Создаем маску для фильтрации строк
mask = (merged_df[rows_to_check] > -48).sum(axis=1) >= 2

# Фильтруем строки, чтобы сохранить 'Monitored Object' и 'Mean RSL'
filtered_df = merged_df[~mask].copy()  # Явное создание копии

# Восстанавливаем значения NaN вместо заполненных значений с использованием .loc
filtered_df.loc[:, rows_to_check] = filtered_df.loc[:, rows_to_check].replace(-9999, np.nan)

filtered_df = filtered_df.drop(columns='Mean TSL', axis=1)

book = load_workbook(rtn_repot_path)
if 'Sheet2' or 'Sheet3' in book.sheetnames:
    del book['Sheet2']
    del book['Sheet3']
book.save(rtn_repot_path)

part_for_drop = [
    '-ODU-1(RTNRF-1)-RTNRF:1', '-MXXI4B-1(IF)-RTNRF:1', '-DMD4-1(IF1)-RTNRF:1', '-DMD4-2(IF2)-RTNRF:1', '-MODU-2(RTNRF-2)-RTNRF:1', '-MODU-1(RTNRF-1)-RTNRF:2',
    '-MODU-2(RTNRF-2)-RTNRF:2', '-MODU-1(RTNRF-1)-RTNRF:1', '-DMD4-1(NM1500-NM1127_A)-RTNRF:1', '-DMD4-2(NM1500-NM1127_B)-RTNRF:1'
]
# Сохраняем данные в Excel
with pd.ExcelWriter(rtn_repot_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    tsl_pivot.to_excel(writer, sheet_name='TSL PIVOT')
    rsl_pivot.to_excel(writer, sheet_name='RSL PIVOT')
    filtered_df.to_excel(writer, sheet_name='RSL & TSL MEAN', index=False)