#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_test_word

def test_columns_generation():
    """Тест для проверки работы генерации в 2 и 3 колонки"""
    
    # Создаем тестовые данные
    test_questions = []
    for i in range(1, 21):  # 20 вопросов
        question = {
            'question_text': f'Вопрос номер {i}. Это тестовый вопрос для проверки работы колонок в Word документе.',
            'options': [
                f'Вариант ответа A для вопроса {i}',
                f'Вариант ответа B для вопроса {i}', 
                f'Вариант ответа C для вопроса {i}',
                f'Вариант ответа D для вопроса {i}'
            ],
            'correct_answer': (i % 4) + 1
        }
        test_questions.append(question)
    
    # Тест 1: Одна колонка (должно работать)
    print("Тестируем 1 колонку...")
    variants_1_col = [{
        'variant_number': 1,
        'questions': test_questions[:10]
    }]
    
    try:
        result_1_col = create_test_word(variants_1_col, 'test_output', columns=1)
        print(f"✓ 1 колонка: успешно создан {result_1_col}")
    except Exception as e:
        print(f"✗ 1 колонка: ошибка - {e}")
    
    # Тест 2: Две колонки
    print("\nТестируем 2 колонки...")
    variants_2_col = [{
        'variant_number': 2,
        'questions': test_questions[:15]
    }]
    
    try:
        result_2_col = create_test_word(variants_2_col, 'test_output', columns=2)
        print(f"✓ 2 колонки: успешно создан {result_2_col}")
    except Exception as e:
        print(f"✗ 2 колонки: ошибка - {e}")
    
    # Тест 3: Три колонки
    print("\nТестируем 3 колонки...")
    variants_3_col = [{
        'variant_number': 3,
        'questions': test_questions
    }]
    
    try:
        result_3_col = create_test_word(variants_3_col, 'test_output', columns=3)
        print(f"✓ 3 колонки: успешно создан {result_3_col}")
    except Exception as e:
        print(f"✗ 3 колонки: ошибка - {e}")
    
    # Тест 4: Смешанный тест - несколько вариантов с разным количеством колонок
    print("\nТестируем смешанный документ...")
    mixed_variants = [
        {
            'variant_number': 1,
            'questions': test_questions[:8]
        },
        {
            'variant_number': 2, 
            'questions': test_questions[:12]
        }
    ]
    
    try:
        result_mixed = create_test_word(mixed_variants, 'test_output', columns=2)
        print(f"✓ Смешанный тест (2 колонки): успешно создан {result_mixed}")
    except Exception as e:
        print(f"✗ Смешанный тест: ошибка - {e}")
    
    print("\nТестирование завершено!")

if __name__ == '__main__':
    test_columns_generation()