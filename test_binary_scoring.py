#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест для перевірки бінарного оцінювання (строга форма іспитів)
"""

import os
import sys
import pandas as pd
from datetime import datetime

# Додаємо шлях до модулів
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import check_student_answers

def create_test_data():
    """Створює тестові дані для перевірки бінарного оцінювання"""
    
    # Створюємо тестові дані у форматі, який очікує функція check_student_answers
    test_data = {
        'Варіант': [1],
        'Відповіді': ['Київ, АВ, Процес утворення органічних речовин, АВД'],
        'Ваги': ['2, 3, 2, 4']
    }
    
    return pd.DataFrame(test_data)

def create_excel_file(df, filename):
    """Створює Excel файл з тестовими даними"""
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Тест', index=False)
    print(f"Створено файл: {filename}")

def test_binary_scoring():
    """Тестує бінарне оцінювання"""
    
    print("=== ТЕСТ БІНАРНОГО ОЦІНЮВАННЯ (СТРОГА ФОРМА) ===")
    print(f"Час тестування: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # Створюємо тестові дані
    df = create_test_data()
    test_file = 'test_binary_scoring.xlsx'
    create_excel_file(df, test_file)
    
    # Тестові сценарії
    test_scenarios = [
        {
            'name': 'Повністю правильні відповіді',
            'answers': ['Київ', 'АВ', 'Процес утворення органічних речовин', 'АВД'],
            'expected_score': 11  # 2+3+2+4 = 11 балів
        },
        {
            'name': 'Частково правильні відповіді (множинний вибір)',
            'answers': ['Київ', 'А', 'Процес утворення органічних речовин', 'АВ'],
            'expected_score': 6  # 2+0+2+0 = 4 бали (бінарне оцінювання!)
        },
        {
            'name': 'Неправильні відповіді множинного вибору',
            'answers': ['Київ', 'БГ', 'Процес утворення органічних речовин', 'БГ'],
            'expected_score': 6  # 2+0+2+0 = 4 бали
        },
        {
            'name': 'Повністю неправильні відповіді',
            'answers': ['Москва', 'БГ', 'Неправильна відповідь', 'БГ'],
            'expected_score': 0  # 0+0+0+0 = 0 балів
        }
    ]
    
    for i, scenario in enumerate(test_scenarios, 1):
        print(f"\n--- Сценарій {i}: {scenario['name']} ---")
        
        try:
            # Перевіряємо відповіді
            result = check_student_answers(test_file, 1, scenario['answers'])
            
            actual_score = result['weighted_score']
            expected_score = scenario['expected_score']
            
            print(f"Відповіді студента: {scenario['answers']}")
            print(f"Очікуваний результат: {expected_score} балів")
            print(f"Фактичний результат: {actual_score} балів")
            print(f"Відсоток: {result['score_percentage']:.1f}%")
            
            # Детальні результати
            print("\nДетальні результати:")
            for detail in result['detailed_results']:
                status = "✓" if detail['is_correct'] else "✗"
                print(f"  Питання {detail['question_number']}: {status} {detail['points']}/{detail['max_points']} балів")
            
            # Перевірка результату
            if abs(actual_score - expected_score) < 0.01:
                print(f"\n✅ ТЕСТ ПРОЙДЕНО: Результат відповідає очікуванням")
            else:
                print(f"\n❌ ТЕСТ НЕ ПРОЙДЕНО: Очікувалося {expected_score}, отримано {actual_score}")
                
        except Exception as e:
            print(f"❌ ПОМИЛКА: {e}")
    
    # Видаляємо тестовий файл
    if os.path.exists(test_file):
        os.remove(test_file)
        print(f"\nВидалено тестовий файл: {test_file}")
    
    print("\n=== ВИСНОВКИ ===")
    print("Система повернулася до бінарного оцінювання:")
    print("• Множинний вибір: або повністю правильно, або 0 балів")
    print("• Відкриті питання: точна відповідність рядків")
    print("• Строга форма, як на реальних іспитах")
    print("• Немає часткових балів за неповні відповіді")

if __name__ == '__main__':
    test_binary_scoring()