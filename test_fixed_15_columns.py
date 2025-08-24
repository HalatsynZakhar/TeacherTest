#!/usr/bin/env python3
"""
Тест для проверки новой логики таблицы ответов:
- Всегда 15 элементов в строке
- Последняя строка дополняется пустыми ячейками для выравнивания ширины
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_test_word
from datetime import datetime

def create_test_variant(num_questions, variant_num=1):
    """Создает тестовый вариант с заданным количеством вопросов"""
    questions = []
    for i in range(num_questions):
        question = {
            'question_text': f'Вопрос {i+1}?',
            'options': [
                f'Вариант A для вопроса {i+1}',
                f'Вариант B для вопроса {i+1}',
                f'Вариант C для вопроса {i+1}',
                f'Вариант D для вопроса {i+1}'
            ],
            'correct_answer': 'A'
        }
        questions.append(question)
    
    variant = {
        'variant_number': variant_num,
        'questions': questions
    }
    return variant

def test_fixed_15_columns():
    """Тестирует новую логику с фиксированными 15 колонками"""
    print("Тестирование новой логики таблицы ответов с фиксированными 15 колонками...")
    
    # Тестовые случаи: разное количество вопросов
    test_cases = [
        (10, "10 вопросов - 1 строка с 5 пустыми ячейками"),
        (15, "15 вопросов - 1 полная строка"),
        (20, "20 вопросов - 2 строки: 15 + 5 с 10 пустыми"),
        (67, "67 вопросов - 5 строк: 4×15 + 7 с 8 пустыми"),
        (30, "30 вопросов - 2 полные строки"),
        (47, "47 вопросов - 4 строки: 3×15 + 2 с 13 пустыми")
    ]
    
    for num_questions, description in test_cases:
        print(f"\nТестирование: {description}")
        
        # Создаем тестовый вариант
        variant = create_test_variant(num_questions)
        variants = [variant]
        
        # Создаем Word документ
        output_dir = "."
        try:
            word_file = create_test_word(variants, output_dir, columns=2)
            print(f"✓ Создан файл: test_{num_questions}_questions_fixed15.docx")
            
            # Переименовываем файл для удобства
            import shutil
            new_name = f"test_{num_questions}_questions_fixed15.docx"
            if os.path.exists(word_file):
                shutil.move(word_file, new_name)
                print(f"✓ Файл переименован в: {new_name}")
            
            # Рассчитываем ожидаемое количество строк
            expected_rows = (num_questions + 14) // 15  # Округление вверх
            empty_cells_in_last_row = (15 - (num_questions % 15)) % 15
            
            print(f"  - Общее количество вопросов: {num_questions}")
            print(f"  - Ожидаемое количество строк таблиц: {expected_rows}")
            print(f"  - Пустых ячеек в последней строке: {empty_cells_in_last_row}")
            
        except Exception as e:
            print(f"✗ Ошибка при создании документа: {e}")
    
    print("\n=== Тестирование завершено ===")
    print("Проверьте созданные Word документы:")
    print("- Каждая строка таблицы должна содержать ровно 15 колонок")
    print("- Последняя строка должна быть дополнена пустыми ячейками")
    print("- Все таблицы должны быть растянуты на всю ширину страницы")
    print("- Заголовок варианта должен быть центрирован вне колонок")
    print("- Таблица ответов должна идти сразу после последнего вопроса")

if __name__ == "__main__":
    test_fixed_15_columns()