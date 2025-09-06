#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест для перевірки виправлення кольорового позначення правильних відповідей,
які не обрав учень (мають бути червоними, а не зеленими)
"""

import pandas as pd
import os
from core.processor import generate_test_variants, create_check_result_word

# Тестові дані з питанням, що має кілька варіантів відповідей
data = {
    'question_number': [1],
    'question': ['Оберіть усі парні числа з наведених варіантів.'],
    'correct_answer': ['АВДЖИ'],  # Правильні відповіді: А, В, Д, Ж, И
    'weight': [4],
    'option_1': ['28'],   # А - правильно
    'option_2': ['15'],   # Б - неправильно
    'option_3': ['80'],   # В - правильно
    'option_4': ['31'],   # Г - неправильно
    'option_5': ['106'],  # Д - правильно
    'option_6': ['99'],   # Е - неправильно
    'option_7': ['44'],   # Ж - правильно
    'option_8': ['57'],   # З - неправильно
    'option_9': ['12'],   # И - правильно
    'task_type': ['ТК']
}

df = pd.DataFrame(data)
df['is_test_question'] = df['task_type'] != 'ВП'

print("=== Тест виправлення кольорового позначення ===\n")

try:
    # Генеруємо варіант
    variants = generate_test_variants(df, num_variants=1)
    variant = variants[0]
    
    print(f"Питання: {variant['questions'][0]['question_text']}")
    print(f"Правильна відповідь: {variant['questions'][0]['correct_answer']}")
    print(f"Варіанти відповідей: {variant['questions'][0]['options']}")
    print()
    
    # Створюємо директорію для результатів
    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)
    
    # Тестуємо випадок, коли учень обрав тільки частину правильних відповідей
    student_answer = 'БВ'  # Учень обрав Б (неправильно) та В (правильно)
    description = 'Учень обрав Б (неправильно) та В (правильно)'
    
    print(f"\n=== {description} ===")
    print(f"Відповідь учня: {student_answer}")
    
    # Створюємо check_result
    check_result = {
        'variant_number': 1,
        'total_questions': 1,
        'correct_answers': 0,  # Неповна відповідь = неправильно
        'score_percentage': 0.0,
        'weighted_score': 0.0,
        'max_score': 4.0,
        'student_info': {
            'class': '5-А',
            'full_name': f'Тест_Кольорів_{student_answer}'
        },
        'detailed_results': [{
            'question_number': 1,
            'question_text': variant['questions'][0]['question_text'],
            'question_type': 'Тестове',
            'student_answer': student_answer,
            'correct_answer': variant['questions'][0]['correct_answer'],
            'is_correct': False,  # Неповна відповідь
            'weight': 4,
            'points': 0.0,
            'max_points': 4.0,
            'is_test_question': True,
            'question_options': variant['questions'][0]['options']
        }]
    }
    
    # Створюємо звіт
    result_path = create_check_result_word(check_result, output_dir)
    print(f"Звіт створено: {result_path}")
    
    # Показуємо, які позначення мають бути у звіті після виправлення
    print("\nОчікувані позначення після виправлення:")
    ukrainian_letters = ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж', 'З', 'И']
    correct_letters = set('АВДЖИ')
    student_letters = set(student_answer)
    
    for i, option in enumerate(variant['questions'][0]['options'], 1):
        letter = ukrainian_letters[i-1]
        status = []
        color = ""
        
        if letter in student_letters and letter in correct_letters:
            status.append("✓ Учень відповів - ПРАВИЛЬНО")
            color = "ЗЕЛЕНИЙ"
        elif letter in student_letters:
            status.append("✗ Учень відповів - НЕПРАВИЛЬНО")
            color = "ЧЕРВОНИЙ"
        elif letter in correct_letters:
            status.append("✓ Правильна відповідь")
            color = "ЧЕРВОНИЙ (ВИПРАВЛЕНО!)"  # Тепер має бути червоним!
        
        status_text = f" ({', '.join(status)})" if status else ""
        color_text = f" [{color}]" if color else ""
        print(f"  {i}. {option}{status_text}{color_text}")
    
    print("\n✅ Тест завершено. Перевірте Word документ:")
    print(f"   - Варіанти А, Д, Ж, И (правильні, але не обрані) мають бути ЧЕРВОНИМИ")
    print(f"   - Варіант В (правильний і обраний) має бути ЗЕЛЕНИМ")
    print(f"   - Варіант Б (неправильний, але обраний) має бути ЧЕРВОНИМ")
    print(f"   - Інші варіанти без позначень")
    
except Exception as e:
    print(f"Помилка: {e}")
    import traceback
    traceback.print_exc()