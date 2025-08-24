#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_test_word

def test_visual_columns():
    """Тест для визуальной проверки работы колонок"""
    
    # Создаем много вопросов для лучшей визуализации колонок
    test_questions = []
    for i in range(1, 31):  # 30 вопросов
        question = {
            'question_text': f'Вопрос {i}. Какой правильный ответ на этот вопрос?',
            'options': [
                f'Первый вариант ответа {i}',
                f'Второй вариант ответа {i}', 
                f'Третий вариант ответа {i}',
                f'Четвертый вариант ответа {i}'
            ],
            'correct_answer': (i % 4) + 1
        }
        test_questions.append(question)
    
    # Тест с 2 колонками
    print("Создаем тест с 2 колонками (30 вопросов)...")
    variants_2_col = [{
        'variant_number': 1,
        'questions': test_questions
    }]
    
    try:
        result_2_col = create_test_word(variants_2_col, 'test_visual_output', columns=2)
        print(f"✓ Документ с 2 колонками создан: {result_2_col}")
        print("Откройте документ и проверьте, что вопросы расположены в 2 колонки")
    except Exception as e:
        print(f"✗ Ошибка при создании документа с 2 колонками: {e}")
    
    # Тест с 3 колонками
    print("\nСоздаем тест с 3 колонками (30 вопросов)...")
    variants_3_col = [{
        'variant_number': 2,
        'questions': test_questions
    }]
    
    try:
        result_3_col = create_test_word(variants_3_col, 'test_visual_output', columns=3)
        print(f"✓ Документ с 3 колонками создан: {result_3_col}")
        print("Откройте документ и проверьте, что вопросы расположены в 3 колонки")
    except Exception as e:
        print(f"✗ Ошибка при создании документа с 3 колонками: {e}")
    
    print("\nТестирование завершено! Проверьте созданные документы в папке test_visual_output")

if __name__ == '__main__':
    test_visual_columns()