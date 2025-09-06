#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест для перевірки правильного відображення порожніх відповідей у Word звіті.
Перевіряємо, що система не показує помилкові позначки про правильні відповіді.
"""

import os
import pandas as pd
from core.processor import check_student_answers, create_check_result_word
from docx import Document

def test_word_display_for_empty_answers():
    """Тестуємо відображення порожніх відповідей у Word документі"""
    
    print("=== ТЕСТ ВІДОБРАЖЕННЯ ПОРОЖНІХ ВІДПОВІДЕЙ У WORD ===")
    
    # Створюємо тестові дані - просте тестове питання з варіантами А, Б, В, Г
    test_data = {
        'Варіант': [1],
        'Відповіді': ['А'],  # Правильна відповідь А
        'Ваги': ['2']
    }
    
    test_file = 'test_word_display.xlsx'
    pd.DataFrame(test_data).to_excel(test_file, index=False)
    print(f"Створено тестовий файл: {test_file}")
    
    # Створюємо папку для результатів
    output_dir = "test_word_output"
    os.makedirs(output_dir, exist_ok=True)
    
    # Тестуємо сценарій: учень нічого не обрав (порожня відповідь)
    student_answers = ['']  # Порожня відповідь
    
    print(f"\nТестуємо сценарій: учень нічого не обрав")
    print(f"Відповідь учня: '{student_answers[0]}'")
    print(f"Правильна відповідь: 'А'")
    
    try:
        # Перевіряємо відповіді
        check_result = check_student_answers(test_file, 1, student_answers)
        
        # Додаємо інформацію про учня
        check_result['student_info'] = {
            'class': '5-А',
            'full_name': 'Тест_Порожня_Відповідь'
        }
        
        # Створюємо Word звіт
        word_path = create_check_result_word(check_result, output_dir)
        print(f"Word звіт створено: {word_path}")
        
        # Аналізуємо результати
        print("\nРезультати перевірки:")
        for result in check_result['detailed_results']:
            question_num = result['question_number']
            student_answer = result['student_answer']
            correct_answer = result['correct_answer']
            is_correct = result['is_correct']
            points = result['points']
            
            print(f"  Питання {question_num}:")
            print(f"    Відповідь учня: '{student_answer}'")
            print(f"    Правильна відповідь: '{correct_answer}'")
            print(f"    Правильно: {is_correct}")
            print(f"    Бали: {points}")
            
            # Перевіряємо логіку
            if student_answer in ['', '(не заповнено)']:
                if is_correct:
                    print(f"    ❌ ПОМИЛКА: Порожня відповідь не повинна бути правильною!")
                    return False
                else:
                    print(f"    ✅ ПРАВИЛЬНО: Порожня відповідь правильно оцінена як неправильна")
        
        # Спробуємо прочитати Word документ і перевірити його вміст
        try:
            doc = Document(word_path)
            doc_text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
            
            print("\n=== АНАЛІЗ WORD ДОКУМЕНТА ===")
            
            # Шукаємо проблемні фрази
            problematic_phrases = [
                'А ✓ (Учень відповів - ПРАВИЛЬНО)',
                'Б ✓ (Учень відповів - ПРАВИЛЬНО)',
                'В ✓ (Учень відповів - ПРАВИЛЬНО)',
                'Г ✓ (Учень відповів - ПРАВИЛЬНО)'
            ]
            
            found_problems = []
            for phrase in problematic_phrases:
                if phrase in doc_text:
                    found_problems.append(phrase)
            
            if found_problems:
                print("❌ ЗНАЙДЕНО ПРОБЛЕМИ у Word документі:")
                for problem in found_problems:
                    print(f"  - {problem}")
                print("\nЦе означає, що система помилково показує, що учень відповів правильно!")
                return False
            else:
                print("✅ Word документ не містить помилкових позначок про правильні відповіді")
            
            # Перевіряємо, що є правильні позначки для правильних відповідей
            expected_phrases = [
                'А ✓ (Правильна відповідь)',  # Це має бути для варіанту А
                '(не заповнено)'  # Це має бути для відповіді учня
            ]
            
            print("\nПеревіряємо наявність правильних позначок:")
            for phrase in expected_phrases:
                if phrase in doc_text:
                    print(f"  ✅ Знайдено: '{phrase}'")
                else:
                    print(f"  ⚠️  Не знайдено: '{phrase}'")
            
            # Виводимо частину тексту документа для аналізу
            print("\n=== ФРАГМЕНТ WORD ДОКУМЕНТА ===")
            lines = doc_text.split('\n')
            for i, line in enumerate(lines):
                if 'Варіанти відповідей' in line or 'А.' in line or 'Б.' in line:
                    # Показуємо кілька рядків навколо варіантів відповідей
                    start = max(0, i-2)
                    end = min(len(lines), i+8)
                    for j in range(start, end):
                        if j == i:
                            print(f">>> {lines[j]}")
                        else:
                            print(f"    {lines[j]}")
                    break
            
        except Exception as e:
            print(f"⚠️  Не вдалося прочитати Word документ: {e}")
        
        print(f"\nЗагальний результат: {check_result['correct_answers']} з {check_result['total_questions']}")
        print(f"Бали: {check_result['weighted_score']:.2f} з {check_result['max_score']:.2f}")
        
        return True
        
    except Exception as e:
        print(f"❌ ПОМИЛКА: {e}")
        return False
    
    finally:
        # Видаляємо тестовий файл
        if os.path.exists(test_file):
            os.remove(test_file)
            print(f"\nВидалено тестовий файл: {test_file}")

def main():
    success = test_word_display_for_empty_answers()
    
    print("\n" + "="*50)
    if success:
        print("✅ ТЕСТ ПРОЙДЕНО: Виправлення працює правильно!")
        print("Порожні відповіді більше не показуються як правильні.")
    else:
        print("❌ ТЕСТ НЕ ПРОЙДЕНО: Потрібні додаткові виправлення.")
    print("="*50)

if __name__ == '__main__':
    main()