#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестування виправлень:
1. Валідація послідовності варіантів відповідей
2. Обробка складних дробів
3. Видалення зайвих дужок
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import process_math_formulas, read_test_excel, create_test_word
import pandas as pd

def test_math_processing():
    """Тестування обробки математичних формул"""
    print("Тестування обробки математичних формул:")
    
    # Тест 1: Складні дроби
    complex_fraction = "\\frac{\\frac{5}{6}}{□(□1/□3)}"
    processed = process_math_formulas(complex_fraction)
    print(f"Складний дріб: {complex_fraction} -> {processed}")
    
    # Тест 2: Зайві дужки
    unnecessary_brackets = "(□(□3/□8))+x=1"
    processed = process_math_formulas(unnecessary_brackets)
    print(f"Зайві дужки: {unnecessary_brackets} -> {processed}")
    
    # Тест 3: Звичайні дроби
    simple_fraction = "□(□4/□7)"
    processed = process_math_formulas(simple_fraction)
    print(f"Звичайний дріб: {simple_fraction} -> {processed}")
    
    # Тест 4: Комбінований вираз
    combined = "x^{2}+(□(□3/□8))+\\frac{\\frac{5}{6}}{□(□1/□3)}=0"
    processed = process_math_formulas(combined)
    print(f"Комбінований: {combined} -> {processed}")

def test_option_validation():
    """Тестуємо валідацію послідовності варіантів відповідей"""
    print("\n=== Тестування валідації варіантів відповідей ===")
    
    # Створюємо тестовий Excel файл з пропуском у варіантах
    # Перший рядок - інструкції (буде пропущений)
    # Далі дані без заголовків (як очікує функція)
    test_data = [
        ['Інструкції', 'для', 'тесту', '', '', '', ''],  # Рядок інструкцій
        [1, 'Тестове питання', 3, 1, 'Варіант 1', None, 'Варіант 3']  # Пропуск у другому варіанті
    ]
    
    df = pd.DataFrame(test_data)
    test_file = 'test_validation.xlsx'
    
    try:
        # Записуємо тестовий файл без заголовків
        df.to_excel(test_file, index=False, header=False)
        
        # Спробуємо прочитати файл - повинна виникнути помилка
        try:
            read_test_excel(test_file)
            print("❌ Валідація НЕ спрацювала - файл з пропусками був прийнятий")
        except ValueError as e:
            if "пропусками у варіантах відповідей" in str(e):
                print("✅ Валідація спрацювала правильно:", str(e)[:100] + "...")
            else:
                print("❌ Валідація спрацювала, але з неочікуваною помилкою:", str(e))
        
    finally:
        # Видаляємо тестовий файл
        if os.path.exists(test_file):
            os.remove(test_file)

def create_test_document():
    """Створення тестового документа з математичними формулами"""
    print("\nСтворення тестового документа...")
    
    # Тестові дані без пропусків
    test_data = {
        'Питання': [
            'Знайдіть число, обернене до числа (□(□4/□7))',
            'Обчисліть \\frac{\\frac{5}{6}}{□(□1/□3)}',
            'Розв\'яжіть рівняння (□(□3/□8))+x=1'
        ],
        'Правильна відповідь': [1, 1, 2],
        'Вага': [1, 1, 1],
        'Варіант 1': ['(□(□7/□4))', '\\frac{5}{2}', 'x=\\frac{5}{8}'],
        'Варіант 2': ['(-□(□4/□7))', '\\frac{5}{18}', 'x=\\frac{8}{5}'],
        'Варіант 3': ['1', '2', 'x=\\frac{3}{8}'],
        'Варіант 4': ['0', '\\frac{15}{6}', 'x=1']
    }
    
    df = pd.DataFrame(test_data)
    test_excel_path = 'test_math_fixes.xlsx'
    df.to_excel(test_excel_path, index=False)
    
    try:
        # Читаємо та обробляємо Excel файл
        from core.processor import generate_test_variants
        df_read = read_test_excel(test_excel_path)
        variants = generate_test_variants(df_read, num_variants=1)
        
        # Створюємо Word документ
        word_path = create_test_word(variants, '.', input_file_name='test_math_fixes')
        print(f"Документ створено: {word_path}")
        
    except Exception as e:
        print(f"Помилка створення документа: {e}")
    finally:
        # Видаляємо тестовий Excel файл
        if os.path.exists(test_excel_path):
            os.remove(test_excel_path)

if __name__ == "__main__":
    test_math_processing()
    test_option_validation()
    create_test_document()
    print("\nТестування завершено!")