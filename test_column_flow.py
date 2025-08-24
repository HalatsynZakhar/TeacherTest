#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_test_word

def test_column_flow():
    """Тест правильного перетекания колонок в Word документе"""
    
    # Создаем много вопросов для тестирования перетекания
    test_variants = [{
        'variant_number': 1,
        'questions': [
            {
                'question_text': f'Вопрос номер {i}: Какой язык программирования используется для веб-разработки?',
                'options': ['Python', 'JavaScript', 'C++', 'Java'],
                'correct_answer': 'B'
            } for i in range(1, 21)  # 20 вопросов для тестирования
        ]
    }]
    
    print("Тестирование перетекания колонок...")
    
    # Создаем Word тест с 3 колонками
    try:
        test_word_path = create_test_word(test_variants, "test_column_flow_reports", columns=3)
        print(f"✅ Word тест с 3 колонками создан: {test_word_path}")
    except Exception as e:
        print(f"❌ Ошибка создания Word теста с 3 колонками: {e}")
    
    # Создаем Word тест с 2 колонками
    try:
        test_word_path = create_test_word(test_variants, "test_column_flow_reports", columns=2)
        print(f"✅ Word тест с 2 колонками создан: {test_word_path}")
    except Exception as e:
        print(f"❌ Ошибка создания Word теста с 2 колонками: {e}")
    
    print("\nПроверьте созданные файлы:")
    print("1. Вопросы должны перетекать из первой колонки во вторую, из второй в третью")
    print("2. При переходе на новую страницу вопросы должны начинаться с первой колонки")
    print("3. Нумерация должна быть сквозной (1, 2, 3... 20)")
    print("4. Не должно быть пустых блоков или неравномерного распределения")

if __name__ == "__main__":
    test_column_flow()