import sys
sys.path.append('.')
from core.processor import read_test_excel, generate_test_variants
import pandas as pd

print("=== Тестирование обработки числовых данных ===")

# Читаем Excel файл
print("\n1. Чтение Excel файла:")
df = read_test_excel('test_numbers.xlsx')
print(df)
print(f"\nТипы данных:\n{df.dtypes}")

# Проверяем конкретные значения
print("\n2. Проверка конкретных значений:")
for col in df.columns:
    if 'Вариант' in col or col == 'Правильный ответ':
        print(f"{col}: {df[col].tolist()} (тип: {type(df[col].iloc[0])})")

# Генерируем варианты теста
print("\n3. Генерация вариантов теста:")
variants = generate_test_variants(df, num_variants=1)
print(f"Количество вариантов: {len(variants)}")

if variants:
    variant = variants[0]
    questions = variant['questions']
    print(f"\nПервый вариант содержит {len(questions)} вопросов:")
    for i, question in enumerate(questions[:2]):  # Показываем первые 2 вопроса
        print(f"\nВопрос {i+1}: {question['question_text']}")
        print(f"Правильный ответ: {question['correct_answer']} (тип: {type(question['correct_answer'])})")
        if 'options' in question and question['options']:
            print("Варианты ответов:")
            for j, option in enumerate(question['options']):
                print(f"  {j+1}) {option} (тип: {type(option)})")