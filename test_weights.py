from core.processor import read_test_excel, generate_test_variants, create_excel_answer_key
import pandas as pd

# Читаем тестовый файл с разными весами
df = read_test_excel('temp/test_weights.xlsx')
print('Прочитанные данные:')
print(df)
print('\nВеса:')
print(df['weight'])

# Генерируем варианты
variants = generate_test_variants(df, 2)
print('\nСгенерированные варианты:')
for v in variants:
    weights = [q['weight'] for q in v['questions']]
    print(f'Вариант {v["variant_number"]}: веса = {weights}')

# Создаем ключ
key_file = create_excel_answer_key(variants, 'temp', 'test_weights')
print(f'\nСоздан ключ: {key_file}')

# Проверяем содержимое ключа
key_df = pd.read_excel(key_file)
print('\nСодержимое ключа:')
print(key_df)
print('\nВеса в ключе:')
print(key_df['Ваги'])