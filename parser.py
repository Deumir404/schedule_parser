import pandas as pd
from openpyxl import load_workbook
import os
from glob import glob

folder_path = "."  # текущая папка, можно заменить на нужную
pattern = "*.xlsx"

# Находим самый последний изменённый файл
files = glob(os.path.join(folder_path, pattern))
if not files:
    raise FileNotFoundError("Excel-файлы не найдены в указанной папке")

latest_file = max(files, key=os.path.getmtime)
print(f"Используется файл: {latest_file}")

# Загружаем последний файл
df = pd.read_excel(latest_file, sheet_name="Колледж ВятГУ")
df = df.iloc[22:]

# Название нужной группы
target_group = "ИСПк-402-52-00"

# Колонка времени
time_col = "Unnamed: 7" 

# Находим колонку с нужной группой
target_col = None
for col in df.columns:
    if df[col].astype(str).str.contains(target_group, na=False).any():
        target_col = col
        break

if target_col and time_col:
    target_index = df.columns.get_loc(target_col)
    group_cols = df.columns[target_index:target_index+4]
    result = pd.concat([df[[time_col]], df[group_cols]], axis=1)

    # --- Добавляем строки с днем недели перед первой встречей ---
    days = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
    new_rows = []
    current_day = 0

    for idx, row in result.iterrows():
        time_value = str(row[time_col]).strip()
        # Считаем, что встреча начинается в 8:20 и далее - начало нового дня
        if time_value == "8.20-9.50" and current_day < len(days):
            # Вставляем строку с названием дня
            new_rows.append([days[current_day]] + [""]*(len(result.columns)-1))
            current_day += 1
        new_rows.append(row.tolist())

    # Создаем новый DataFrame
    result_with_days = pd.DataFrame(new_rows, columns=result.columns)

    # Сохраняем в Excel
    file_out = f"./schedule/{target_group.replace(' ', '_')}.xlsx"
    result_with_days.to_excel(file_out, index=False, engine="openpyxl")

    # Настраиваем ширину колонок
    wb = load_workbook(file_out)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(file_out)
    print(f"Расписание для группы {target_group} сохранено в {file_out} с днями недели и автошириной")

else:
    print("Не удалось найти все нужные столбцы.")
