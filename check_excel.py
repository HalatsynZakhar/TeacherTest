import pandas as pd

# Читаємо Excel файл без пропуску рядків
df = pd.read_excel('test_numbers.xlsx')
print('Структура файлу (з заголовками):')
print(df.head())
print('\nКолонки:')
print(df.columns.tolist())
print('\nПерші 3 рядки:')
for i in range(min(3, len(df))):
    print(f'Рядок {i}: {df.iloc[i].tolist()}')

print('\n' + '='*50)
print('Читаємо з пропуском першого рядка (як у коді):')
df_skip = pd.read_excel('test_numbers.xlsx', header=None, skiprows=1)
print('Структура:')
print(df_skip.head())
print('\nПерші 3 рядки:')
for i in range(min(3, len(df_skip))):
    print(f'Рядок {i}: {df_skip.iloc[i].tolist()}')