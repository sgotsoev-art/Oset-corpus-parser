import pandas as pd

# Загружаем файл
INPUT_FILE = "final_results.xlsx"
OUTPUT_FILE = "final_results.xlsx"

# Читаем данные
df = pd.read_excel(INPUT_FILE, sheet_name=0)

# Получаем список всех уникальных фразеологизмов
all_phrases = df['Фразеологизм'].unique()

# Получаем список всех подкорпусов
all_subcorpora = df['Подкорпус'].unique()

# Оставляем только строки с НЕ пустым контекстом (реальные совпадения)
df_real = df[df['Контекст'].notna() & (df['Контекст'] != '')]

# Считаем частоты для реальных совпадений
freq_counts = df_real.groupby(['Фразеологизм', 'Подкорпус']).size().reset_index(name='частота')

# Создаём пустую таблицу со всеми комбинациями (фразеологизм × подкорпус)
all_combinations = pd.MultiIndex.from_product([all_phrases, all_subcorpora], names=['Фразеологизм', 'Подкорпус'])
freq_full = pd.DataFrame(index=all_combinations).reset_index()

# Объединяем с реальными частотами (там где нет совпадений — будет NaN)
freq_full = freq_full.merge(freq_counts, on=['Фразеологизм', 'Подкорпус'], how='left')

# Заменяем NaN на 0
freq_full['частота'] = freq_full['частота'].fillna(0).astype(int)

# Превращаем в таблицу: строки — фразеологизмы, столбцы — подкорпуса
freq_table = freq_full.pivot(index='Фразеологизм', columns='Подкорпус', values='частота')

# Переименовываем колонки
freq_table.columns = [f"Частота в {col}" for col in freq_table.columns]

# Добавляем колонку "Всего"
freq_table['Всего по всем подкорпусам'] = freq_table.sum(axis=1)

# Сохраняем в тот же Excel-файл, на новый лист
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    freq_table.to_excel(writer, sheet_name='Частота')

print(f"Частотная таблица добавлена на лист 'Частота' в {OUTPUT_FILE}")
print(f"Всего фразеологизмов: {len(freq_table)}")
print(f"Всего подкорпусов: {len(all_subcorpora)}")