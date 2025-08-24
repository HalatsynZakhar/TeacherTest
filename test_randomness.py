import sys
sys.path.append('.')
from core.processor import read_test_excel, generate_test_variants
import pandas as pd

print("=== Тестирование случайности генерации ===\n")

# Читаем Excel файл
df = read_test_excel('test_numbers.xlsx')
print(f"Загружено {len(df)} вопросов\n")

# Генерируем несколько наборов вариантов для проверки случайности
print("Проверка случайности генерации (первые 3 вопроса во втором варианте):")
print("-" * 60)

for run in range(1, 6):  # 5 запусков
    print(f"\nЗапуск {run}:")
    variants = generate_test_variants(df, num_variants=3)
    
    if len(variants) >= 2:
        variant2 = variants[1]  # Второй вариант
        questions = variant2['questions']
        
        print(f"  Вариант 2 содержит {len(questions)} вопросов:")
        for i, question in enumerate(questions[:3]):  # Первые 3 вопроса
            question_text = question['question_text'][:50] + "..." if len(question['question_text']) > 50 else question['question_text']
            print(f"    {i+1}. {question_text}")
            
            if 'options' in question and question['options']:
                options_preview = ", ".join([str(opt) for opt in question['options'][:3]])
                if len(question['options']) > 3:
                    options_preview += "..."
                print(f"       Варианты: {options_preview}")

print("\n" + "=" * 60)
print("Если генерация случайна, порядок вопросов и ответов должен отличаться между запусками.")
print("Если вы видите одинаковые последовательности - проблема не решена.")