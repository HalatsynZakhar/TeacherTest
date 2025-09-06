#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест для перевірки виправлення проблеми з порожніми відповідями.
Перевіряємо, що коли учень нічого не обирав, система не показує,
що він відповів правильно на перший варіант.
"""

import os
import pandas as pd
from core.processor import check_student_answers, create_check_result_word

def create_test_data():
    """Створює тестові дані для перевірки обробки порожніх відповідей"""
    
    # Створюємо тестові дані у форматі, який очікує функція check_student_answers
    test_data = {
        'Варіант': [1],
        'Відповіді': ['А, БВ, Київ'],  # Тестове, множинний вибір, відкрите
        'Ваги': ['2, 3, 2']
    }
    
    return pd.DataFrame(test_data)

def test_empty_answer_handling():
    """Тестуємо обробку порожніх відповідей"""
    
    print("=== ТЕСТ ОБРОБКИ ПОРОЖНІХ ВІДПОВІДЕЙ ===")
    
    # Створюємо тестовий файл
    test_file = 'test_empty_answers.xlsx'
    test_data = create_test_data()
    test_data.to_excel(test_file, index=False)
    print(f"Створено тестовий файл: {test_file}")
    
    # Створюємо папку для результатів
    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)
    
    # Тестові сценарії з порожніми відповідями
    test_scenarios = [
        {
            'name': 'Повністю порожні відповіді',
            'answers': ['', '', ''],
            'description': 'Учень нічого не відповів на жодне питання'
        },
        {
            'name': 'Частково порожні відповіді',
            'answers': ['', 'Б', ''],
            'description': 'Учень відповів тільки на друге питання'
        },
        {
            'name': 'Порожня відповідь на тестове питання',
            'answers': ['', 'БВ', 'Київ'],
            'description': 'Учень не відповів на перше тестове питання'
        }
    ]
    
    for scenario in test_scenarios:
        print(f"\n--- {scenario['name']} ---")
        print(f"Опис: {scenario['description']}")
        print(f"Відповіді учня: {scenario['answers']}")
        
        try:
            # Перевіряємо відповіді
            check_result = check_student_answers(test_file, 1, scenario['answers'])
            
            # Додаємо інформацію про учня
            check_result['student_info'] = {
                'class': '5-А',
                'full_name': f'Тест_{scenario["name"].replace(" ", "_")}'
            }
            
            # Створюємо Word звіт
            word_path = create_check_result_word(check_result, output_dir)
            print(f"Word звіт створено: {word_path}")
            
            # Аналізуємо результати
            print("\nРезультати перевірки:")
            for result in check_result['detailed_results']:
                question_num = result['question_number']
                student_answer = result['student_answer']
                is_correct = result['is_correct']
                points = result['points']
                
                print(f"  Питання {question_num}:")
                print(f"    Відповідь учня: '{student_answer}'")
                print(f"    Правильно: {is_correct}")
                print(f"    Бали: {points}")
                
                # Перевіряємо, що порожні відповіді не вважаються правильними
                if student_answer in ['', '(не заповнено)'] and is_correct:
                    print(f"    ❌ ПОМИЛКА: Порожня відповідь не повинна бути правильною!")
                elif student_answer in ['', '(не заповнено)'] and not is_correct:
                    print(f"    ✅ ПРАВИЛЬНО: Порожня відповідь правильно оцінена як неправильна")
            
            print(f"\nЗагальний результат: {check_result['correct_answers']} з {check_result['total_questions']}")
            print(f"Бали: {check_result['weighted_score']:.2f} з {check_result['max_score']:.2f}")
            
        except Exception as e:
            print(f"❌ ПОМИЛКА: {e}")
    
    # Видаляємо тестовий файл
    if os.path.exists(test_file):
        os.remove(test_file)
        print(f"\nВидалено тестовий файл: {test_file}")
    
    print("\n=== ВИСНОВКИ ===")
    print("Перевірено обробку порожніх відповідей:")
    print("• Порожні відповіді не повинні показуватися як правильні")
    print("• У Word звіті не повинно бути позначок про правильні відповіді для порожніх полів")
    print("• Система повинна коректно відображати '(не заповнено)' для порожніх відповідей")

if __name__ == '__main__':
    test_empty_answer_handling()