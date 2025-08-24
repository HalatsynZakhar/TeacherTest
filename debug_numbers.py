import pandas as pd
from core.processor import read_test_excel, generate_test_variants

print("=== Отладка форматирования чисел ===")

# Проверим файл test_numbers.xlsx
file_path = "test_numbers.xlsx"
print(f"\nЧтение файла: {file_path}")

try:
    df = read_test_excel(file_path)
    print("\nДанные из Excel:")
    print(df.to_string())
    
    print("\nТипы данных:")
    print(df.dtypes)
    
    print("\nПроверка конкретных значений в столбцах с вариантами ответов:")
    for col in df.columns:
        if col.startswith('option_'):
            print(f"\n{col}:")
            for idx, val in df[col].items():
                if pd.notna(val):
                    print(f"  Строка {idx}: '{val}' (тип: {type(val)})")
    
    # Генерируем варианты
    print("\n=== Генерация вариантов ===")
    variants = generate_test_variants(df, 1)
    
    if variants:
        variant = variants[0]
        questions = variant['questions']
        print(f"\nВариант содержит {len(questions)} вопросов:")
        
        for i, question in enumerate(questions[:3]):  # Показываем первые 3 вопроса
            print(f"\nВопрос {i+1}: {question['question_text']}")
            if 'options' in question and question['options']:
                print("Варианты ответов:")
                for j, option in enumerate(question['options']):
                    print(f"  {j+1}) '{option}' (тип: {type(option)})")
                    
except Exception as e:
    print(f"Ошибка: {e}")
    import traceback
    traceback.print_exc()